import copy
import os

from dotenv import load_dotenv
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (Attachment, Disposition, FileContent,
                                   FileName, FileType, Mail)

from input_output import GSpreadAuthentication
from problem_solving import KandGProblem
from read_df import csvtodf, makeical


def main(path_input=None, dir_output=None, direct_in=False, local=False):
    if direct_in:
        if local:
            # Google API秘密鍵のパスを入力
            json_keyfile_name = "hoge/key_file_name.json"
            # https://docs.google.com/spreadsheets/d/xxx/....のxxx部分を入力
            spreadsheet_key = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
            spreadsheet = GSpreadAuthentication(json_keyfile_name, spreadsheet_key)
        else:
            spreadsheet = GSpreadAuthentication()

        df_kagisime, df_gomisute = spreadsheet.InputSchedule()
    else:
        path_input = input("Please enter path of schedule csv file:\n") if path_input is None else path_input
        csvdata = csvtodf(path_input)
        df_kagisime, df_gomisute = csvdata.df_input()

    model = KandGProblem(df_kagisime, df_gomisute)
    df_output = model.solve()
    print(df_output)
    print(df_output.index[0])
    for name, series in df_output.iteritems():
        print(name, series[5])
    making_ical = makeical(df_output)

    if direct_in:
        spreadsheet.UpdateSpreadsheet(df_output)

        mail_content_temp = spreadsheet.GetDecidedMailTemplate()
        mail_content = copy.deepcopy(mail_content_temp)
        print('mail_content', mail_content)
        print('mail_content_temp', mail_content_temp, '\n\n')

        
    if local:
        df_output.to_csv(os.path.join(dir_output, csvdata.yyyymm + ' 配置.csv'), encoding='utf_8_sig')
        dir_output = "./example" if dir_output is None else dir_output
        if not (os.path.isdir(dir_output)):
            os.mkdir(dir_output)

    # .icsファイルを各メンバーごとに作成
    for to_mail, fname in zip(spreadsheet.to_mails, spreadsheet.FName):
        member = to_mail[1]
        encoded_file = making_ical.convert(member, fname)  # ゴミ捨てに登録されている全員のicsファイルを作成
        mail_content['PlainTextContent'] = mail_content_temp['PlainTextContent'].replace('recipient', member)

        if direct_in:
            send_mail(encoded_file, to_mail, mail_content, spreadsheet.yyyymm+'_'+fname)
        else:
            with open(os.path.join(dir_output, csvdata.yyyymm+'_'+fname + '.ics'), mode='wb') as f:
                f.write(encoded_file)


def send_mail(encoded_file, to_mail, mail_content, icsfilename):

    attachedfile = Attachment(
        FileContent(encoded_file),
        FileName(icsfilename+'.ics'),
        FileType('application/ics'),
        Disposition('attachment')
    )

    to_emails = list((to_mail, tuple(os.environ['FROM_MAIL'].split())))

    message = Mail(
        from_email=tuple(os.environ['FROM_MAIL'].split()),
        to_emails=to_emails,
        subject=mail_content['Title'],
        plain_text_content=mail_content['PlainTextContent'],
        is_multiple=True)
    message.attachment = attachedfile
    try:
        sg = SendGridAPIClient(os.environ.get('SENDGRID_API_KEY'))
        print(to_emails)
        response = sg.send(message)
        print(response.status_code)
        print(response.body)
        print(response.headers)
    except Exception as e:
        print(e.message)


if __name__ == '__main__':
    # load_dotenv(verbose=True)
    # dotenv_path = './.env'
    # load_dotenv(dotenv_path)
    # print(os.environ['EXAMPLE'])
    path_input = "./example/GSS_test - 202111 予定表.csv"  # Type path of schedule table csv file downloaded from spreadsheet unless direct_in=True.
    dir_output = "./example"  # Type path of output file(.ical, 予定表.csv).
    main(path_input=None, dir_output=None, direct_in=True, local=False)

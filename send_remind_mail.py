import gspread
from oauth2client.service_account import ServiceAccountCredentials 
import pandas as pd
import os
import datetime
from dotenv import load_dotenv
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail

def main(local=False):
    if local:
        load_dotenv(verbose=True)
        dotenv_path='./.env'
        load_dotenv(dotenv_path)
        print(os.environ['EXAMPLE'])

    dt_now = datetime.datetime.now()
    workbook=gaccount_auth()
    #Select sheet named yyyymm 予定表 and get its URL.
    intyyyymm=dt_now.year*100+dt_now.month+1 if dt_now.month!=12 else (dt_now.year+1)*100+1
    yyyymm=str(intyyyymm)
    worksheetname=yyyymm+" 予定表"
    worksheet = workbook.worksheet(worksheetname)
    schedule_url='https://docs.google.com/spreadsheets/d/'+os.environ['SHEET_SPREADSHEET_KEY_KG']+'/edit#gid='+str(worksheet.id)
    #Select sheet named "MailTemp" and get templates of remind mail.
    MailTempWorksheet = workbook.worksheet('MailTemp')
    RemindMailTempCell=MailTempWorksheet.find('RemindMailTemp')
    remindmail_content_temp=MailTempWorksheet.cell(RemindMailTempCell.row, RemindMailTempCell.col+1).value
    RemindMailTitleTempCell=MailTempWorksheet.find('RemindMailTitleTemp')
    remindmail_Title_temp=MailTempWorksheet.cell(RemindMailTitleTempCell.row, RemindMailTitleTempCell.col+1).value

    Sender=tuple(os.environ['FROM_MAIL'].split())
    NextMonth=dt_now.month+1 if dt_now.month!=12 else 1
    Subject=remindmail_Title_temp.replace('NextMonth', str(NextMonth))
    Content=remindmail_content_temp.replace('ScheduleSheetURL', schedule_url).replace('Sender', Sender[1]).replace('NextMonth', str(NextMonth)).replace('Recipient', '平山研・荒井研M1の皆様')
    
    #Make mailing list from spreadsheet.
    MailingListWorksheet = workbook.worksheet('MailingList')
    cell_list = MailingListWorksheet.get_all_values()
    df=pd.DataFrame(cell_list[1:][:], columns=cell_list[0])
    to_mails=[(mail, name) for mail, name in zip(df['Mail'], df['Name'])]

    message = Mail(
        from_email=Sender,
        to_emails=to_mails,
        subject=Subject,
        plain_text_content=Content)
    try:
        sg = SendGridAPIClient(os.environ.get('SENDGRID_API_KEY'))
        response = sg.send(message)
        print(response.status_code)
        print(response.body)
        print(response.headers)
    except Exception as e:
        print(e.message)

 
def gaccount_auth():
    """Set authentication info of google api. Credential data is called from environmental variables.
    Set env. var. of credential data of google api key in advance.
    """
    
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']

    pk = "-----BEGIN PRIVATE KEY-----{key}-----END PRIVATE KEY-----\n".format(key=os.environ['SHEET_PRIVATE_KEY_STR'].replace('\\n', '\n'))

    #Dict. object. Calls authentication info from env. variables.
    credential = {
                    "type": "service_account",
                    "project_id": os.environ['SHEET_PROJECT_ID'],
                    "private_key_id": os.environ['SHEET_PRIVATE_KEY_ID'],
                    "private_key": pk,
                    "client_email": os.environ['SHEET_CLIENT_EMAIL'],
                    "client_id": os.environ['SHEET_CLIENT_ID'],
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                    "client_x509_cert_url":  os.environ['SHEET_CLIENT_X509_CERT_URL']
                }

    credentials =  ServiceAccountCredentials.from_json_keyfile_dict(credential, scope)
    #Log in Google API using OAuth2 authentication information.
    gc =gspread.authorize(credentials)

    #Open spreadsheet using spreadsheet key.
    SPREADSHEET_KEY = os.environ['SHEET_SPREADSHEET_KEY_KG']
    workbook = gc.open_by_key(SPREADSHEET_KEY)
    return workbook

def gaccount_auth_local(json_keyfile_name, spreadsheet_key):
    """Set authentication info of google api. Credential data is loaded from json file.
    Validate Google drive and spreadsheet API and set authentication info in advance. See "https://developers.google.com/workspace/guides/create-project".
    DO NOT USE WITH CLOUD SERVICE OR GITHUB.
    """
        #Set spreadsheet API and google drive API.
    SCOPES = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

    #Set authentication info.
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile_name, SCOPES)

    #Log in Google API using OAuth2 authentication information。
    gc =gspread.authorize(creds)

    #Open spreadsheet using spreadsheet key.
    SPREADSHEET_KEY = spreadsheet_key
    workbook = gc.open_by_key(SPREADSHEET_KEY)
    return workbook

if __name__ == '__main__':
    path_input="./example/GSS_test - 202111 予定表.csv"#Type path of schedule table csv file downloaded from spreadsheet unless direct_in=True.
    dir_output="./example"#Type path of output file(.ical, 予定表.csv).
    main()

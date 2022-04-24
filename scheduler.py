import datetime

import member_decision
import send_remind_mail

dt_now = datetime.datetime.now()

if dt_now.day == 20:
    send_remind_mail.main(local=False)
elif dt_now.day == 25:
    member_decision.main(path_input=None, dir_output=None, direct_in=True, local=False)

""" load_dotenv(verbose=True)
dotenv_path='./.env'
load_dotenv(dotenv_path)
 """

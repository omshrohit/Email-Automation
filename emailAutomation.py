import imaplib
import email
import pandas as pd
import os,re
import time
def process_email():
        EMAIL = "Your Gmail id"
        PASSWORD = "your app password with 16 characters"
        EXCEL_FILE = "patients.xlsx"

        # create file if not exists 
        old_data = None
        if os.path.exists(EXCEL_FILE):
            try:
                old_data = pd.read_excel(EXCEL_FILE)
            except PermissionError:
                print("close your excel file other wise you can't read and write")
                
        else:
            create_columns = pd.DataFrame(columns=["Email_ID","Name","Age","Disease","Hospital"])
            create_columns.to_excel(EXCEL_FILE,index=False)

        # login

        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(EMAIL,PASSWORD)
        print("Login Success")

        #inbox
        mail.select("inbox")

        status,msg =  mail.search(None,'(UNSEEN SUBJECT "Patient Report")')
        email_msg_ids = msg[0].split()


        if email_msg_ids:
            for email_id in email_msg_ids:
                    try:
                        # old data frame
                        old_df = pd.read_excel(EXCEL_FILE)
                        str_email_id = email_id.decode()
                        if str_email_id in old_df["Email_ID"].astype(str).values:
                            print("get duplicate record")
                            mail.store(email_id,'+FLAGS','\\Seen')
                            continue
                        status,msg = mail.fetch(email_id,"(RFC822)")
                        row_data = msg[0][1]
                        email_msg = email.message_from_bytes(row_data)

                        print(email_msg["Subject"],email_msg["From"],sep="\n")

                        body  = None
                        if email_msg.is_multipart():
                            for part in email_msg.walk():
                                if part.get_content_type() == "text/plain":
                                    body = part.get_payload(decode=True).decode()
                                    break
                        else:
                            body = email_msg.get_payload(decode=True).decode()

                        
                        name_match = re.search(r"Patient Name:\s*(.*)",body)
                        match_age = re.search(r"Age:\s*(\d+)",body)
                        match_disease = re.search(r"Disease:\s*(.*)",body)
                        match_hospital = re.search(r"Hospital:\s*(.*)",body)

                        if not name_match or not match_age:
                            print("Name and Age is required")
                            continue
                        
                        name = name_match.group(1).strip() if name_match else None
                        age = match_age.group(1).strip() if match_age else None
                        disease = match_disease.group(1).strip() if match_disease else None
                        hospital = match_hospital.group(1).strip() if match_hospital else None
                        print(name,age,disease,hospital)

                        new_df = pd.DataFrame([{
                            "Email_ID":str_email_id,
                            "Name":name,
                            "Age":age,
                            "Disease":disease,
                            "Hospital":hospital
                        }])
                        updated_df = pd.concat([old_df,new_df],ignore_index=True)
                        updated_df.to_excel(EXCEL_FILE,index=False)
                        print("new record added")

                        mail.store(email_id,'+FLAGS','\\Seen')
                    except PermissionError as e:
                        print(f"Cclose your excel file other wise you can't read and write {e}")
                        break
                    except Exception as e:
                        print(f"email processing error {e}")
                        continue
        else:
            print("Did't get any new mail!")

        mail.logout()
        print("logout")

while True:
    try:
        process_email()
    except Exception as e:
        print(e)
    time.sleep(120)

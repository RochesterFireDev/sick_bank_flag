import os
import json
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

def load_flags(path):
    """
    Loads a json file, or creates an empty json file, to store data for members flagged (yellow or red) vfor their sick-leave bank 

    Parameters:
    path (type): path to the json files

    Returns:
    json: a json file containing flagged members    
    """
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        data = []
        with open(path, "w", encoding="utf-8") as f:
            json.dump([], f)

    flags = []
    for x in data:
        if isinstance(x, dict):
            emp  = int(x.get("employee_id"))
            name = str(x.get("name", ""))
            date_yellow = x.get("date_yellow")
            date_red    = x.get("date_red")
        else:  # legacy list of IDs
            emp, name, date_yellow, date_red = int(x), "", None, None

        flags.append({
            "employee_id": emp,
            "name": name,
            "date_yellow": date_yellow,
            "date_red": date_red
        })
    return flags

def save_flags(path, flags):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(flags, f, indent=2)


def send_email(body, file_path, test=True):
    """
    Sends an email specifically for the sick and injured report payroll, case manager, and their chain of command. 

    Parameters:
    body (string): An html string of the email body text
    file_path (str): the filepath of the sick_and_injured.xlsx report 
    test (bool): If True, then it sends to developers and outputs report locally, if False, pipeline 
                runs in production: sends to stakeholders and exports report to production folder.

    Returns:
    email: sends an email 
    """
    SMTP_SERVER = "smtp.cor.local"
    SMTP_PORT = 25
    SMTP_USER = "rfdutilsvc@cityofrochester.gov"
    FROM_EMAIL = SMTP_USER

    if test:
        to_emails=["willis.sontheimer@cityofrochester.gov"
                  # ,"Daniel.Curran@cityofrochester.gov"
                   ]
    else:
        to_emails = ["Erica.Torres@CityofRochester.Gov"
                    , "rfd-payroll@CityofRochester.Gov"
                    , "Tracy.Brown@CityofRochester.Gov"
                    , "Michael.Dobbertin@CityofRochester.Gov"
                    , "Arthur.Kucewicz@CityofRochester.Gov"
                    ]
    
    message = MIMEMultipart()
    message["From"] = FROM_EMAIL
    message["To"] = ", ".join(to_emails)
    message["Subject"] = f"Sick Bank Flag — {datetime.now():%Y-%m-%d}"
    message["X-Priority"] = "1"            # 1 (High) .. 5 (Low)
    message["X-MSMail-Priority"] = "High"
    message["Importance"] = "High"
    message.attach(MIMEText(body, "html"))

    # Attach file
    if isinstance(file_path, str) and os.path.exists(file_path):
        with open(file_path, "rb") as file:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition",
                        f'attachment; filename="{os.path.basename(file_path)}"')
        message.attach(part)

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.sendmail(SMTP_USER, to_emails, message.as_string())
            print(f"Email sent from {FROM_EMAIL} to {', '.join(to_emails)} with attachment: {file_path}")
    except Exception as e:
        print(f"Error sending email: {e}")
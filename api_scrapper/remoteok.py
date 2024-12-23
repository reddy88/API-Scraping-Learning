import requests
import xlwt
from xlwt import Workbook, easyxf
import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate, COMMASPACE

#https://myaccount.google.com/lesssecureapps

BASE_URL = "https://remoteok.io/api"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
REQUEST_HEADERS = {
    'User-Agent': USER_AGENT,
    'Accept_Language': 'en-US,en;q=0.5'
}

def get_job_postings():
    # Get the data from the API
    response = requests.get(url=BASE_URL, headers=REQUEST_HEADERS)
    return response.json()

def filter_data(data, columns_to_keep):
    # Keep only specified columns
    filtered_data = []
    for job in data:
        filtered_job = {col: job.get(col, "") for col in columns_to_keep}
        filtered_data.append(filtered_job)
    return filtered_data

def save_jobs_to_excel(data):
    # Create a new Excel file
    wb = Workbook()
    sheet = wb.add_sheet('Jobs')
    header_style = easyxf('font: bold 1; align: horiz center;')
    headers = data[0].keys()
    for i, header in enumerate(headers):
        sheet.write(0, i, header, header_style)
    for i, data in enumerate(data):
        for j, value in enumerate(data.values()):
            sheet.write(i+1, j, value)
            
    # Save the Excel file
    wb.save('remoteok_jobs.xls')

def send_emails(send_to,send_from,subject,text,files=None):
    assert isinstance(send_to, list)
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject
    msg.attach(MIMEText(text))
    for f in files or []:
        with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        msg.attach(part)

        smtp = smtplib.SMTP('smtp.gmail.com', 587)
        smtp.starttls()
        smtp.login(send_from, 'password')
        smtp.sendmail(send_from, send_to, msg.as_string())
        smtp.close()

if __name__ == "__main__":
    json = get_job_postings()[1:]
    columns_to_keep = [
        'date', 'company', 'company_logo', 'position', 'tags',
        'description', 'location', 'salary_min', 'salary_max', 'apply_url'
    ]
    cleaned_data = filter_data(json, columns_to_keep)
    print(save_jobs_to_excel(cleaned_data))
    send_emails(['xyz@gmail.com','abc@gmail.com','Jobs Posting', 'Please, find the attached file.', ['remoteok_jobs.xls']])
import requests
import xlwt
from xlwt import Workbook
import smtplib
from os.path import basename

from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate, COMMASPACE

BASE_URL='https://remoteok.com/api'
U_AGENT ='Mozilla/5.0 (X11; Linux x86_64; rv:128.0) Gecko/20100101 Firefox/128.0'
R_HEADERS = {
    'User-Agent': U_AGENT,
    'Accept-Language': 'en-US, en;q=0.5',
}

def get_jobs():
    res = requests.get(url=BASE_URL, headers=R_HEADERS)
    return res.json()


if __name__ == "__main__":
    jobs = get_jobs()[1]
    print(jobs)
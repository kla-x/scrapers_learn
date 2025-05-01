import requests
import xlwt
from xlwt import Workbook
import csv
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

def save_spreadsheet(jobs):
    wb = Workbook()
    j_sheet = wb.add_sheet('jobs')
    bheaders = list(jobs[0].keys())
   
    for i in range(0, len(bheaders)):
        j_sheet.write(0,i,bheaders[i])
    
    for i in range(0, len(jobs)):
        values = list(jobs[i].values())
        for n in range(0, len(values)):
            j_sheet.write(i+1,n,values[n])
    wb.save('jobs_sheet.xls')


def save_csv(jobs):
   
    bheaders = list(jobs[1].keys())
    brows = []
    for i in range(0, len(jobs)):
        brows.append(list(jobs[i].values()))
    
    with open('jobs_lisings.csv', 'w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file, quoting=csv.QUOTE_ALL)
        writer.writerow(bheaders)
        writer.writerows(brows)

if __name__ == "__main__":
    jobs = get_jobs()[1:]
    save_spreadsheet(jobs)  
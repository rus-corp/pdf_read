import pdfminer.high_level
import re
from pprint import pprint
from openpyxl import load_workbook


with open('pdftext.txt', 'r', encoding='utf-8') as f:
    file = f.readlines()


def contacts():
    phone_number = re.findall(r'\+7\s*\D*\d+\D*\d+\D*\d+\D?\d+', str(file))

    email_contact = re.findall(r'\S*@\w+.\w+[m|u]', str(file))
    print(len(phone_number))
    print(len(email_contact)) 
    # total_list = phone_number + email_contact


         

    fn = 'data.xlsx'
    wb = load_workbook(fn)
    ws = wb['data1']
    for phone in phone_number:
        ws.append([phone])
    for email in email_contact:
        ws.append([email])

    

    wb.save('data1.xlsx')
    wb.close()


    # with open ('contacts.txt', 'w', encoding='utf-8') as f:
    #     for line in total_list:
    #         f.write(line + '\n')



contacts()






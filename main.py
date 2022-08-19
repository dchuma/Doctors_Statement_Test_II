from pathlib import Path
import pandas as pd
# import win32com.client as win32
import math
import numpy as np
import os
DOCTORS_EMAIL_ACCOUNT ='doctorsaccounts@nbihosp.org'

statement_date = '31.07.2022'
PATH = r"G:\My Drive\PROGRAMMING_LANGUAGES\MyPythonExercises\Doctors_Samples\AUGUST2022_STATEMENTS"
EXCEL_FILE_PATH = os.path.join(PATH,'DRS.OUTPUT_ASAT04082022copy.xlsx')
OUTPUT = os.path.join(PATH, 'OUTPUT_DRS.STATEMENT_TEST_I.xlsx')
ATTACHMENT_DIR = os.path.realpath('Attachments')
# ATTACHMENT_DIR.mkdir(exist_ok=True)
data = pd.read_excel(EXCEL_FILE_PATH, sheet_name="Data")
type(PATH)
data.replace(np.nan,0)
billing = data['GROSS_BILLING']
payment = data['GROSS_PAID']

print(type(billing))
print(type(payment))
# convert float to int
billing.astype(int)
payment.fillna(0).astype(int)
data['BALANCE'] = billing.fillna(0).astype(int) - payment.fillna(0).astype(int)
conditions = [
    (payment > billing),
    (payment < billing),
    (payment == billing),
    (billing == 0),
    (payment == 0)
] ##

choices = ['OVER-PAID','PARTLY-PAID','BILL-PAID','BILL-WITHDRAWN','PENDING']

data['STATUS'] = np.select(conditions,choices, default = 'PENDING')
data.columns
df = data.drop(['DOCUMENT_NO','DOCUMENT_NO.2','CODE','CODE PAID','DESCRIPTION','DOCTORS.PAID','PV REF. PAID','INVOICE NO. PAID'], axis=1).replace(np.nan,)
df.columns
df.style.format({'GROSS_BILLING': '{0:,.2f}'})
df.style.format({'GROSS_PAID': '{0:,.2f}'})
df.style.format({'BALANCE': '{0:,.2f}'})
df.style.format({"DATE": lambda t: t.strftime("%d/%m/%Y")})
column_name = "DOCTORS"
unique_values = df[column_name].unique()
unique_values
df.query("DOCTORS == 'DR. ABWAO HENRY O'").head(10)
df.to_excel(OUTPUT, index=False)
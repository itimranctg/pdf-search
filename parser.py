import sys
import re
import pdfplumber
import pandas as pd
from openpyxl import Workbook

pdf_file = sys.argv[1]

rows = []

with pdfplumber.open(pdf_file) as pdf:

    for page in pdf.pages:

        text = page.extract_text()

        if not text:
            continue

        blocks = re.split(r'(?=\d{4}\.\s*নাম:)', text)

        for block in blocks:

            serial = ""
            name = ""
            voter = ""
            father = ""
            mother = ""
            job = ""
            dob = ""
            address = ""

            m = re.search(r'(\d{4})\.\s*নাম:\s*(.*)', block)
            if m:
                serial = m.group(1)
                name = m.group(2).split("\n")[0]

            m = re.search(r'ভোটার নং:\s*([0-9০-৯]+)', block)
            if m:
                voter = m.group(1)

            m = re.search(r'পিতা:\s*(.*)', block)
            if m:
                father = m.group(1).split("\n")[0]

            m = re.search(r'মাতা:\s*(.*)', block)
            if m:
                mother = m.group(1).split("\n")[0]

            m = re.search(r'পেশা:\s*(.*?),\s*জন্ম', block)
            if m:
                job = m.group(1)

            m = re.search(r'জন্ম.*?:\s*([0-9/০-৯]+)', block)
            if m:
                dob = m.group(1)

            m = re.search(r'ঠিকানা:\s*(.*)', block)
            if m:
                address = m.group(1).replace("\n"," ")

            if voter:
                rows.append([
                    serial,
                    name,
                    voter,
                    father,
                    mother,
                    job,
                    dob,
                    address
                ])

df = pd.DataFrame(rows, columns=[
    "ক্রমিক",
    "নাম",
    "ভোটার নং",
    "পিতা",
    "মাতা",
    "পেশা",
    "জন্মতারিখ",
    "ঠিকানা"
])

import os
os.makedirs("output", exist_ok=True)

df.to_excel("output/result.xlsx", index=False)

print("Done")

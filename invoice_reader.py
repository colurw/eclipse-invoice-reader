# for Windows only...?
# winget install -e --id UB-Mannheim.TesseractOCR    <- paste into terminal
# https://ghostscript.com/releases/gsdnld.html       <- download and install appropriate file


import pdfplumber
import os
from dateutil.parser import parse
from datetime import date
import re
import pandas as pd
from pandas import DataFrame
import ocrmypdf


folder = 'raw_invoices'    
# folder = 'raw_invoices_test' 
year = '2023'
all_records = []
failed = []


# read text from pdf/a files 

for filename in os.listdir(folder):
    if filename.endswith('.pdf'): 
        print(filename)
        with pdfplumber.open(f'{folder}/{filename}') as file:          
            text = ''
            for page in file.pages:
                chunk = page.extract_text(x_tolerance=3, 
                                          x_tolerance_ratio=None, 
                                          y_tolerance=3, layout=False, 
                                          x_density=7.25, 
                                          y_density=13, 
                                          line_dir_render=None, 
                                          char_dir_render=None,)
                text += '\n\n'+ chunk   
            print(text) 

            record = {}
            record['id'] = int(''.join(filter(str.isdigit, filename)))                  


            # a.kapllaj

            if 'pllaj' in text.lower():
                record['from'] = 'A. Kapllaj'

                lines = text.split('\n')
                for i, line in enumerate(lines):

                    if 'hoffman' in line.lower() or 'collendale' in line.lower() or 'champness' in line.lower():
                        if 'to' not in record.keys():
                            splitted = line.split(':')
                            splitted = splitted[1].split(', ')
                            for s_line in splitted:
                                if 'hoffman' in s_line.lower() or 'collendale' in s_line.lower() or 'champness' in s_line.lower():
                                    record['to'] = s_line

                    if 'invoice date:' in line.lower():
                        words = line.split(':')
                        inv_date = parse(words[-1]).strftime('%d/%m/%Y')
                        record['date'] = inv_date

                    if line == 'FOR Amount':
                        j = i
                        descr = []
                        while True:
                            if '£' in lines[j+2]:
                                break
                            descr.append(lines[j+1])
                            j += 1
                        record['description'] = ''.join(descr)

                    if 'Barclays Bank' in line:
                        splitted = line.split('£')
                        cost = re.sub('[, ]', '', splitted[-1])
                        record['cost'] = int(float(cost))           


            # contour landscapes

            elif 'contour landscapes' in text.lower():
                record['from'] = 'Contour Landscapes'

                lines = text.split('\n')
                for i, line in enumerate(lines):

                    if 'invoice date:' in line.lower():
                        words = line.split(':')
                        inv_date = parse(words[-1]).strftime('%d/%m/%Y')
                        record['date'] = inv_date

                    if line == 'Works Completed':
                        j = i
                        descr = []
                        while True:
                            descr.append(lines[j+1])
                            if 'Sub Total' in lines[j+2]:
                                break
                            j += 1
                        record['description'] = ' '.join(descr)

                    if 'APPLICATION TOTAL' in line:
                        splitted = line.split('£')
                        cost = re.sub('[, ]', '', splitted[-1])
                        record['cost'] = int(float(cost))


            # DPC

            elif 'discreet pest' in text.lower():
                record['from'] = 'Discrete Pest Control'

                lines = text.split('\n')
                for i, line in enumerate(lines):

                    if 'date' in line.lower():
                        words = line.split(' ')
                        inv_date = parse(words[-1]).strftime('%d/%m/%Y')
                        record['date'] = inv_date

                    if line == 'Description Unit Price Total':
                        j = i
                        descr = []
                        while True:
                            if 'Sub Total' in lines[j+1]:
                                break
                            descr.append(lines[j+1])
                            j += 1
                        record['description'] = ' '.join(descr)

                    if 'Amount' in line:
                        splitted = line.split('£')
                        cost = re.sub('[, ]', '', splitted[-1])
                        record['cost'] = int(float(cost))


            # EMS 

            elif 'expert maintenance solutions' in text.lower():
                record['from'] = 'Expert Maintenance Solutions'

                lines = text.split('\n')
                for i, line in enumerate(lines):

                    if 'invoice date' in line.lower():
                        words = line.split(':')
                        inv_date = words[1].split(' ')
                        record['date'] = inv_date[1]

                    if line == 'Invoice Header':
                        j = i
                        descr = []
                        while True:
                            if 'Items Quantity Unit' in lines[j+1]:
                                break
                            descr.append(lines[j+1])
                            j += 1
                        record['description'] = ' '.join(descr)

                    if 'Total Including VAT' in line:
                        splitted = line.split('£')
                        cost = re.sub('[, ]', '', splitted[-1])
                        record['cost'] = int(float(cost))


            # E.on - type 1 (eg. 3537417)

            elif 'eonnext' in text.lower():
                record['from'] = 'E.on'
                record['type'] = 1

                lines = text.split('\n') 
                for i, line in enumerate(lines):

                    if 'Date issued' in line:
                        inv_date = re.sub('th', '', lines[i+1])
                        inv_date = re.sub('nd', '', inv_date)
                        inv_date = re.sub('st', '', inv_date)
                        inv_date = parse(inv_date).strftime('%d/%m/%Y')
                        record['date'] = inv_date

                    if 'Supply Address: ' in line:
                        record['description'] = line.split(': ')[-1]

                    if 'Charges for Meter' in line:
                        record['meter'] = line.split('Meter')[-1].split(' ')[1]

                    if 'Total Electricity Charges' in line:
                        splitted = line.split(' ')
                        cost = re.sub('[, £]', '', splitted[-1])
                        record['cost'] = int(float(cost))


            # eurosafe

            elif 'eurosafe' in text.lower():
                record['from'] = 'Eurosafe'

                lines = text.split('\n')
                for i, line in enumerate(lines):

                    if 'invoice date' in line.lower():
                        words = line.split(':')
                        inv_date = parse(words[-1], dayfirst=True).strftime('%d/%m/%Y')
                        # inv_date = words[1].split(' ')
                        record['date'] = inv_date

                    if line == 'QTY Description Item Price Total':
                        j = i
                        descr = []
                        while True:
                            if 'Total excl. VAT' in lines[j+1]:
                                break
                            descr.append(lines[j+1])
                            j += 1
                        record['description'] = ' '.join(descr)

                    if 'Invoice Total' in line:
                        splitted = line.split('GBP')
                        cost = re.sub('[, ]', '', splitted[-1])
                        record['cost'] = int(float(cost))      


            # JPD

            elif 'jpd' in text.lower():
                record['from'] = 'JPD Cleaning Services'

                lines = text.split('\n')
                for i, line in enumerate(lines):

                    if 'TERMS' in line:
                        words = line.split(' ')
                        inv_date = parse(str(words[1])).strftime('%d/%m/%Y')
                        record['date'] = inv_date

                    if 'DESCRIPTION QTY RATE AMOUNT' in line:
                        j = i
                        descr = []
                        while True:
                            if 'Please make payment to' in lines[j+1]:
                                break
                            descr.append(lines[j+1])
                            j += 1
                        record['description'] = ' '.join(descr)

                    if 'TOTAL DUE' in line:
                        cost = re.sub('[, £]', '', lines[i-1])
                        record['cost'] = int(float(cost))


            # KGN Pillinger

            elif 'kgnpillinger' in text:
                record['from'] = 'KGN Pillinger'

                lines = text.split('\n')
                for i, line in enumerate(lines):

                    if 'Invoice date' in line:
                        words = line.split(' ')
                        record['date'] = parse(str(words[-1])).strftime('%d/%m/%Y')
                        

                    if line == 'Description Qty Unit Price Nett VAT':
                        record['description'] = str(lines[i+1])

                    if 'invoice total' in line.lower():
                        splitted = line.split()
                        cost = re.sub('[, ]', '', splitted[-1])
                        record['cost'] = int(float(cost)) 


            # MJ Fire

            elif 'mj fire' in text.lower():
                record['from'] = 'MJ Fire Safety'

                lines = text.split('\n')
                for i, line in enumerate(lines):

                    if 'DATE:' in line:
                        words = line.split(' ')
                        inv_date = parse(str(words[1])).strftime('%d/%m/%Y')
                        record['date'] = inv_date

                    if 'Quantity Details' in line:
                        j = i
                        descr = []
                        while True:
                            if 'PLEASE NOTE' in lines[j+1] or 'BACS:' in lines[j+1]:
                                break
                            descr.append(lines[j+1])
                            j += 1
                        record['description'] = ' '.join(descr)

                    if 'Invoice Total' in line:
                        splitted = line.split('£')
                        cost = re.sub('[, £]', '', splitted[-1])
                        record['cost'] = int(float(cost))


            # PLP Fire

            elif 'plp fire' in text.lower():
                record['from'] = 'PLP Fire & Security'

                lines = text.split('\n')
                for i, line in enumerate(lines):

                    if 'Invoice Date' in line:
                        if 'TAX' in line:
                            words = lines[i+1].split('Ltd')
                        else:
                            words = lines[i+2].split('Ltd')
                        inv_date = parse(str(words[0])).strftime('%d/%m/%Y')
                        record['date'] = inv_date

                    if 'Description Quantity' in line:
                        j = i
                        descr = []
                        while True:
                            if 'Subtotal' in lines[j+1]:
                                break
                            descr.append(lines[j+1])
                            j += 1
                        record['description'] = ' '.join(descr)

                    if 'TOTAL GBP' in line:
                        splitted = line.split(' ')
                        cost = re.sub('[, £]', '', splitted[-1])
                        record['cost'] = int(float(cost))


            # Remus

            elif 'VAT Reg No: 568 5406 11' in text:
                record['from'] = 'REMUS'

                lines = text.split('\n')
                for i, line in enumerate(lines):

                    if 'Invoice Date' in line:
                        words = line.split(': ')
                        inv_date = parse(str(words[-1])).strftime('%d/%m/%Y')
                        record['date'] = inv_date

                    if 'Property Description' in line:
                        j = i
                        descr = []
                        while True:
                            if 'Totals' in lines[j+1]:
                                break
                            descr.append(lines[j+1])
                            j += 1
                        record['description'] = ' '.join(descr)

                    if 'Invoice £' in line:
                        splitted = line.split(' ')
                        cost = re.sub('[, £]', '', splitted[-1])
                        record['cost'] = int(float(cost))


            # Schindler

            elif 'schindler' in text.lower():
                record['from'] = "Schindler"

                lines = text.split('\n')
                for i, line in enumerate(lines):

                    if 'Date: ' in line:
                        if 'date' not in record.keys():
                            words = line.split(': ')
                            inv_date = parse(str(words[-1])).strftime('%d/%m/%Y')
                            record['date'] = inv_date

                    if 'job description' in line.lower():
                        j = i
                        descr = []
                        while True:
                            if 'Report No' in lines[j]:
                                break
                            descr.append(lines[j])
                            j += 1
                        record['description'] = ' '.join(descr)
                    
                    if 'ALERT 24/7' in line:
                        record['description'] = '24/7 callout standby'

                    if 'Total to pay' in line:
                        splitted = line.split(' ')
                        cost = re.sub('[, £]', '', splitted[-1])
                        record['cost'] = int(float(cost))


            # SE Controls

            elif 'secontrols' in text.lower():
                record['from'] = "SE Controls"

                lines = text.split('\n')
                for i, line in enumerate(lines):

                    if 'Date' in line:
                        if 'date' not in record.keys():
                            words = line.split(' ')
                            inv_date = parse(str(words[-1])).strftime('%d/%m/%Y')
                            record['date'] = inv_date

                    if 'Qty Description' in line:
                        j = i
                        descr = []
                        while True:
                            if 'Total GBP Excl' in lines[j+1]:
                                break
                            descr.append(lines[j+1])
                            j += 1
                        record['description'] = ' '.join(descr)

                    if 'Total GBP Incl' in line:
                        splitted = line.split(' ')
                        cost = re.sub('[, £]', '', splitted[-1])
                        record['cost'] = int(float(cost))


            # Security Shutters

            elif 'securityshutterslimited' in text.lower():
                record['from'] = "Security Shutters"

                lines = text.split('\n')
                for i, line in enumerate(lines):

                    if 'InvoiceDate' in line:
                        if 'date' not in record.keys():
                            inv_date = parse(lines[i+1]).strftime('%d/%m/%Y')
                            record['date'] = inv_date

                    if 'Description Quantity' in line:
                        j = i
                        descr = []
                        while True:
                            if 'Subtotal' in lines[j+1]:
                                break
                            descr.append(lines[j+1])
                            j += 1
                        record['description'] = ' '.join(descr)

                    if 'TOTALGBP' in line:
                        splitted = line.split(' ')
                        cost = re.sub('[, £]', '', splitted[-1])
                        record['cost'] = int(float(cost))


            # unbloc 

            elif 'Unbloc Drainage' in text:
                record['from'] = 'Unbloc Drainage'

                lines = text.split('\n')
                for i, line in enumerate(lines):

                    if 'hoffman' in line.lower() or 'collendale' in line.lower() or 'champness' in line.lower():
                        splitted = line.split(', ')
                        for s_line in splitted:
                            if 'hoffman' in s_line.lower() or 'collendale' in s_line.lower() or 'champness' in s_line.lower():
                                record['to'] = s_line 

                    if line == 'Date Order Ref. Job No. Account No. Invoice No.':
                        record['date'] = str(lines[i+1][0:10])

                    if line == 'Description of Works':
                        record['description'] = str(lines[i+1])

                    if 'total to pay' in line.lower():
                        splitted = line.split()
                        cost = [s_line for s_line in splitted if '£' in s_line][0]
                        splitted = cost.split('.')
                        cost = [cost for cost in splitted if '£' in cost][0]
                        record['cost'] = int(''.join(filter(str.isdigit, cost)))


            # Verity

            elif 'veritylandcare' in text.lower():
                record['from'] = "Verity Landcare"

                lines = text.split('\n')
                for i, line in enumerate(lines):

                    if 'Tax Date' in line:
                        if 'date' not in record.keys():
                            inv_date = lines[i+2].split()
                            inv_date = parse(' '.join(inv_date[3:6])).strftime('%d/%m/%Y')
                            record['date'] = inv_date

                    if 'Description Quantity' in line:
                        j = i
                        descr = []
                        while True:
                            if 'Subtotal' in lines[j+1]:
                                break
                            descr.append(lines[j+1])
                            j += 1
                        record['description'] = ' '.join(descr)

                    if 'Amount Due GBP' in line:
                        splitted = line.split(' ')
                        cost = re.sub('[, £]', '', splitted[-1])
                        record['cost'] = int(float(cost))


            # WheelieWashers

            elif 'info@commercialbinhygiene' in text.lower():
                record['from'] = "WheelieWashers"

                lines = text.split('\n')
                for i, line in enumerate(lines):

                    if 'Date:' in line:
                        if 'date' not in record.keys():
                            inv_date = parse(lines[i+2]).strftime('%d/%m/%Y')
                            record['date'] = inv_date

                    if 'Quantity Description' in line:
                        j = i
                        descr = []
                        while True:
                            if 'VAT' in lines[j+2]:
                                break
                            descr.append(lines[j+1])
                            j += 1
                        record['description'] = ' '.join(descr)

                    if 'VAT at 20' in line:
                        splitted = lines[i+1].split(' ')
                        cost = re.sub('[, £]', '', splitted[-1])
                        record['cost'] = int(float(cost))



            # convert image-based pdf to a text-based pdf/a

            elif __name__ == '__main__':
                if filename not in os.listdir('ocr_temp'):
                    ocrmypdf.ocr(f'{folder}/{filename}', f'ocr_temp/{filename}', deskew=True, force_ocr=True)

                with pdfplumber.open(f'ocr_temp/{filename}') as file:
                    text = ''
                    for page in file.pages:
                        chunk = page.extract_text(x_tolerance=3, 
                                                x_tolerance_ratio=None, 
                                                y_tolerance=3, layout=False, 
                                                x_density=7.25, 
                                                y_density=13, 
                                                line_dir_render=None, 
                                                char_dir_render=None,)
                        text += '\n\n'+ chunk   
                    print(text)
    

                # fawcetts

                if 'fawcetts' in text.lower():
                    record['from'] = 'Fawcetts Accountants'

                    lines = text.split('\n')
                    for i, line in enumerate(lines):

                        if 'invoice date fawcetts' in line.lower():
                            words = lines[i+2].split(' ')
                            print(words[:3])
                            inv_date = parse(''.join(words[:3]), dayfirst=True).strftime('%d/%m/%Y')
                            record['date'] = inv_date

                        if 'vat number' in line.lower():
                            j = i
                            descr = []
                            while True:
                                if 'Subtotal' in lines[j+2]:
                                    break
                                descr.append(lines[j+2])
                                j += 1
                            record['description'] = ' '.join(descr)

                        if 'TOTAL GBP' in line:
                            splitted = line.split('GBP')
                            cost = re.sub('[, £]', '', splitted[-1])
                            record['cost'] = int(float(cost))   


                # eon - type 2 (eg. 3468272)

                if 'eonnext' in text.lower() and 'your energy account' in text.lower():
                    record['from'] = "E.on"
                    record['type'] = 2
                    lines = text.split('\n')
                    for i, line in enumerate(lines):

                        if 'Bill Reference:' in line:
                            if 'date' not in record.keys():
                                line = re.sub('[)]', '', line)
                                splitted = line.split('(')
                                inv_date = parse(splitted[-1]).strftime('%d/%m/%Y')
                                record['date'] = inv_date

                        if 'Supply Address' in line:
                            descr = line.split(':')
                            record['description'] = descr[-1]

                        if 'Charges for Meter' in line:
                            # if 'meter' not in record.keys():
                            splitted = line.split('Meter ')
                            meter = splitted[-1]
                            # print(meter)
                            meter = meter.split(' ')[0]
                            # print(meter)
                            record['meter'] = meter

                        if 'Total charges for bill' in line:
                            splitted = line.split(' ')
                            cost = re.sub('[, €£]', '', splitted[-1])
                            record['cost'] = int(float(cost))

                
                # eon - type 3  (eg. 3468577)

                elif 'eonnext' in text.lower() and 'invoice' in text.lower():
                    record['from'] = "E.on"
                    record['type'] = 3
                    lines = text.split('\n')
                    for i, line in enumerate(lines):

                        if 'Invoice' in line and 'Tax' not in line:
                            if 'date' not in record.keys():
                                inv_date = re.sub('th', '', lines[i-1])
                                inv_date = re.sub('nd', '', inv_date)
                                inv_date = re.sub('st', '', inv_date)
                                print(inv_date)
                                inv_date = inv_date.split(' ')
                                inv_date = ' '.join(inv_date[-3:])
                                try:
                                    inv_date = parse(inv_date).strftime('%d/%m/%Y')
                                except:
                                    try:
                                        inv_date = parse('01 '+ inv_date[2:]).strftime('%d/%m/%Y')
                                    except:
                                        pass
                                record['date'] = inv_date

                        if 'Supply Address' in line:
                            record['description'] = line.split(':')[-1]

                        if 'Charges for Meter' in line:
                            if 'meter' not in record.keys():
                                splitted = line.split('Meter ')
                                meter = splitted[-1].split(' ')[0]
                                record['meter'] = meter

                        if 'Total charges for bill' in line:
                            splitted = line.split(' ')
                            cost = re.sub('[, €£]', '', splitted[-1])
                            record['cost'] = int(float(cost))

                # Ferndale

                elif 'ferndale' in text.lower():
                    record['from'] = "Ferndale Insurance"

                    lines = text.split('\n')
                    for i, line in enumerate(lines):

                        if 'Itm No.' in line:
                            j = i
                            descr = []
                            while True:
                                if 'Invoice Balance' in lines[j+1]:
                                    break
                                descr.append(lines[j+1])
                                j += 1
                            record['description'] = ' '.join(descr)

                        if 'Invoice Balance' in line:
                            splitted = line.split(' ')
                            cost = re.sub('[, £]', '', splitted[-1])
                            record['cost'] = int(float(cost))

                    date_pattern = "(.*)\d{1,2}\/\d{1,2}\/\d{2,4}"
                    compiled = re.compile(date_pattern)
                    for line in lines:
                        match = compiled.match(line)
                        if match:
                            if 'date' not in record.keys():
                                inv_date = parse(line.split(' ')[-1], dayfirst=True).strftime('%d/%m/%Y')
                                record['date'] = inv_date


                # SSE swalec

                elif 'swalec' in text.lower():
                    record['from'] = "SSE Swalec"

                    lines = text.split('\n') 
                    for i, line in enumerate(lines):

                        if 'Tax point date' in line:
                            inv_date = re.sub('[?]', '', line.split('date')[-1])
                            inv_date = parse(inv_date).strftime('%d/%m/%Y')
                            record['date'] = inv_date

                        if 'Supply to: ' in line:
                            record['description'] = line.split(': ')[-1]

                        if 'Meter Number' in line:
                            record['meter'] = line.split(' ')[-1]

                        if 'Total this invoice' in line:
                            splitted = line.split(' ')
                            cost = re.sub('[, £]', '', splitted[-1])
                            record['cost'] = int(float(cost))


            # unrecognised contractor

            else: 
                failed.append(filename)


            all_records.append(record)
            print(record)


if '_test' not in folder:
    df = DataFrame(all_records)
    print(df.head)
    print(df.tail)

    df.to_pickle('df.pkl')  
    df = pd.read_pickle('df.pkl')
    df.to_excel(f"{year}_invoices.xlsx")

print('unreadable invoices:', failed)
import xlrd, xlwt
import sys
from datetime import datetime

def parseBankRow(row, i):
    record_object = {
        'index': i,
        'serialNo': row[0].value,
        'date1' : row[1].value,
        'date2' : row[2].value,
        'modeOfPayment' : row[3].value,
        'chequeNo' : row[4].value,
        'randNo1': row[5].value,
        'randNo2': row[6].value,
        'amount' : row[7].value,
        'col9' : row[9].value
    }
    return record_object

def parseCashRow(row, i):
    pay_object = {
        'index': i,
        'serialNo': row[0].value,
        'chequeNo': row[1].value,
        'date': row[2].value,
        'amount': row[3].value,
        'entryNo':row[4].value
    }
    return pay_object
entries_needed = []
def findEntry(cash, bank_database):
    for record in bank_database:
	mode = record['modeOfPayment']
	refNo = str(mode).split()[-1]
    	if refNo.isdigit():
        	to_check = refNo
    	else:
        	to_check = record['chequeNo'] 
	#if record['chequeNo'] != '':
         #   print 'here', record['serialNo']
          #  record['chequeNo'] = int(record['chequeNo']) 
        if str(to_check).isdigit():
            to_check = int(to_check) 
        if str(cash['chequeNo']) == to_check and cash['amount'] == record['amount']:
            return record['serialNo']
    entries_needed.append((to_check, cash['amount']))
    return -1

def sheetToDatabase(sheet):
    database = []
    for i in xrange(sheet.nrows):
        row = sheet.row(i)
        record_object = parseBankRow(row, i)
        database.append(record_object)
    return database

def cashSheetToDatabase(sheet):
    database = []
    for i in xrange(sheet.nrows):
        row = sheet.row(i)
        record_object = parseCashRow(row, i)
        database.append(record_object)
    return database

def main():
    if len(sys.argv) < 3:
        print '''
        usage: python %s bank.xlsx cash.xls [OUTFILE.xls]
        OUTFILE.xls if not specified, taken from prompt.
        If OUTFILE.xls is simply '-' then don't write to Excel but
        simple text file to stdout''' % sys.argv[0]
        exit()
    bank_file = sys.argv[1]
    print 'Opening', bank_file, 'x ...'
    book = xlrd.open_workbook(bank_file, on_demand = True)
    sheet = book.sheet_by_index(0)
    print 'Reading data...'
    bank_database = sheetToDatabase(sheet)
    
    cash_file = sys.argv[2]
    print 'Opening', cash_file, 'x ...'
    book = xlrd.open_workbook(cash_file, on_demand = True)
    sheet = book.sheet_by_index(0)
    print 'Reading data...'
    cash_database = cashSheetToDatabase(sheet)
    

    print 'Total Records:', len(bank_database)
    if len(sys.argv) == 4:
        output_file = sys.argv[3]
    else:
        output_file = raw_input('Enter the output file name:')
    if output_file == '-':
        print '\n'.join("%s, %s" % (r['accountNo'], findBalance(r)) for r in database)
    elif output_file:
        print 'Writing to ', output_file, ' ...'
        workbook = xlwt.Workbook()
        worksheet = workbook.add_sheet('Sheet 1')
        for r in cash_database:
            worksheet.write(r['index'], 0, r['serialNo'])
            worksheet.write(r['index'], 1, r['chequeNo'])
            worksheet.write(r['index'], 2, r['date'])
            worksheet.write(r['index'], 3, r['amount'])
            worksheet.write(r['index'], 4, r['entryNo'])
            worksheet.write(r['index'], 5, findEntry(r, bank_database))
        workbook.save(output_file)
        print len(entries_needed)
if __name__ == '__main__':
    main()

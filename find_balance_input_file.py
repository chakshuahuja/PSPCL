import xlrd, xlwt
import sys
from datetime import datetime

def parseRecordRow(row, i):
    record_object = {
        'index': i,
        'serialNo': row[0].value,
        'chequeNo' : row[1].value,
        'date' : row[2].value,
        'amount' : row[3].value,
        'entryNo' : row[4].value,
    }
    return record_object

def sheetToDatabase(sheet):
    database = []
    for i in xrange(sheet.nrows):
        row = sheet.row(i)
        record_object = parseRecordRow(row, i)
            
        if len(database):
            last_record = database.pop()
            if record_object['chequeNo'] == last_record['chequeNo']:
                last_record['amount'] += record_object['amount'] 
                database.append(last_record)
            else:
                database.append(last_record)
                database.append(record_object)
        else:
            database.append(record_object)

    return database

def main():
    if len(sys.argv) < 2:
        print '''
        usage: python %s INFILE.xlsx [OUTFILE.xls]
        OUTFILE.xls if not specified, taken from prompt.
        If OUTFILE.xls is simply '-' then don't write to Excel but
        simple text file to stdout''' % sys.argv[0]
        exit()
    input_file = sys.argv[1]
    print 'Opening', input_file, 'x ...'
    book = xlrd.open_workbook(input_file, on_demand = True)
    sheet = book.sheet_by_index(0)
    print 'Reading data...'
    database = sheetToDatabase(sheet)
    print 'Total Records:', len(database)
    
    if len(sys.argv) == 3:
        output_file = sys.argv[2]
    else:
        output_file = raw_input('Enter the output file name:')
    if output_file == '-':
        print '\n'.join("%s, %s" % (r['accountNo'], findBalance(r)) for r in database)
    elif output_file:
        print 'Writing to ', output_file, ' ...'
        workbook = xlwt.Workbook()
        worksheet = workbook.add_sheet('Sheet 1')
        for r in database:
            worksheet.write(r['index'], 0, r['serialNo'])
            worksheet.write(r['index'], 1, r['chequeNo'])
            worksheet.write(r['index'], 2, r['date'])
            worksheet.write(r['index'], 3, r['amount'])
            worksheet.write(r['index'], 4, r['entryNo'])
        workbook.save(output_file)
if __name__ == '__main__':
    main()



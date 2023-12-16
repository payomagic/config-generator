import re
import xlrd

def replace_words(text, word_dic):
    rc = re.compile('|'.join(map(re.escape, word_dic)))
    def translate(match):
        return word_dic[match.group(0)]
    return rc.sub(translate, text)

workbook = xlrd.open_workbook('Book1.xlsx')
worksheet = workbook.sheet_by_name('Sheet1')
num_rows = worksheet.nrows - 1
curr_row = 1

while curr_row <= num_rows:
    row = worksheet.row(curr_row)

    hostname = row[0].value
    loopback0 = row[1].value
    ip_address = row[2].value
    netmask = row[3].value

    with open('Template1.txt', 'r') as t:
        tempstr = t.read()

    values = {
        '[Hostname]': hostname,
        '[Loopback0]': loopback0,
        '[IP_Address]': ip_address,
        '[Netmask]': netmask,
    }

    outputfile = f"{hostname}_CONFIG.txt"
    output = replace_words(tempstr, values)

    with open(outputfile, 'w') as fout:
        fout.write(output)

    curr_row += 1

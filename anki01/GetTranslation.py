"""
驗證單字是否有拼錯
"""

from googletrans import Translator
import csv
import sys
import xlrd


def readExcel(filename):
    workbook = xlrd.open_workbook(filename)
    sheet = workbook.sheet_by_index(0)
    card_lst = []
    for rowx in range(sheet.nrows):
        cols = sheet.row_values(rowx)
        print(cols)
        card_lst.append(cols)
    return card_lst


ankitxtfile = 'Feb.4.xlsx'

if not len(sys.argv) < 2:
    ankitxtfile = sys.argv[1]

if ankitxtfile.find('xls') > 1:
    pkgname = ankitxtfile.replace('.xlsx', '').replace('.xls', '')
    card_lst = readExcel(ankitxtfile)
else:
    pkgname = ankitxtfile.replace('.txt', '')
    with open(ankitxtfile, 'r', encoding='utf8') as f:
        cardstr = f.read()
    card_lst = list(csv.reader(cardstr.splitlines(), delimiter='\t'))

translator = Translator()

for c in card_lst:
    c.append(translator.translate(c[0], dest='zh-TW').text)

with open('verification_' + pkgname + '.csv', mode='w', encoding='utf-8-sig', newline='') as f:
    csv_writer = csv.writer(f)
    csv_writer.writerows(card_lst)

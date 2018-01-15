import genanki
import os
from gtts import gTTS
import csv
import xlrd
import sys


def getGTTS(word):
    """
    獲得Google翻譯的語音
    :param word:
    :return:
    """
    tts = gTTS(text=word, lang='en', slow=True)
    filename = word.replace('/', '') + ".mp3"
    if not os.path.exists(filename):
        tts.save(filename)
        print('saved: ' + filename)
    else:
        print(filename + ' 已存在')
    return filename


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
# print(card_lst)

#############################

for c in card_lst:
    c.append(getGTTS(c[0]))

# print(card_lst)

###############################################################

qfmt = '<p style="text-align:center;font-family: arial;font-size: 20px;">'
qfmt += '{{Front}}</p>'

afmt = '<p style="text-align:center;font-family: arial;font-size: 20px;">'
afmt += '{{FrontSide}}</p>\n\n<hr id="answer">\n\n'
afmt += '<p style="text-align:center;font-family: arial;font-size: 20px;">'
afmt += '{{Back}}</p>'

my_model = genanki.Model(
    1607392319,
    'Simple Model',
    fields=[
        {'name': 'Front'},
        {'name': 'Back'},
    ],
    templates=[
        {
            'name': 'Card 1',
            'qfmt': qfmt,
            'afmt': afmt,
        },
    ])

my_deck = genanki.Deck(
    2059400110,
    pkgname)

media_lst = []
for card in card_lst:
    my_deck.add_note(genanki.Note(
        model=my_model,
        fields=[card[0] + '[sound:' + card[2] + ']', card[1]]
    ))
    media_lst.append(card[2])
    # print([card[0] + '[sound:' + card[2] + ']', card[1]])

# my_deck.add_note(my_note)
my_package = genanki.Package(my_deck)
my_package.media_files = media_lst
my_package.write_to_file(pkgname + '.apkg')
print(pkgname + '.apkg', 'OK')

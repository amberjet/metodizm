import openpyxl
from docxtpl import DocxTemplate
import pymorphy2
from num2words import num2words
day30 = ['апрель', 'июнь', 'сентябрь', 'ноябрь']
day31 = ['январь', 'март', 'май', 'июль', 'август', 'октябрь', 'декабрь']
PRICEFL = [2874.0, 1724.1, 6896.5]
PRICESZ = [2659.6, 1595.7, 6383.0]
prilmonth = 'октябрь'
prilyear = '2021'
filelink = 'https://docs.google.com/spreadsheets/d/1IedgMGBxmEjWgyokDRXrgwmG7U_9Ym55'
name = 'Методисты_Работа_за_октябрь21.xslx'
morph = pymorphy2.MorphAnalyzer()
m = morph.parse(prilmonth)[0]


def make_file(record):
    # если методист в этот месяце что-то делал
    if record[-1]:
        initial = record[6].split()
        initial = initial[0] + ' ' + initial[1][0] + '. ' + initial[2][0] + '.'
        m2 = morph.parse(record[0])[0]
        # формируем общую часть шаблона
        context = {'dogmonth': m2.inflect({'sing', 'gent'}).word, 'dogyear': record[1], 'prilnum': record[3],
                   'subj': record[4], 'dognum': record[5], 'fio': record[6], 'finit': initial,
                   'prilmonthrp': m.inflect({'sing', 'gent'}).word,
                   'prilmonthpp': m.inflect({'sing', 'loct'}).word,
                   'prilyear': prilyear}
        if prilmonth in day30:
            context['lastday'] = '30'
        elif prilmonth in day31:
            context['lastday'] = '31'
        else:
            context['lastday'] = '28'
        filename = record[6].split()[0]
        allsum = 0
        # часть для физлиц
        if record[2] == 'ФЛ':
            templ_name1 = 'template.docx'
            templ_name2 = 'template_akt.docx'
            PRICE = PRICEFL
        else:
            templ_name1 = 'templateSZ.docx'
            templ_name2 = 'templateSZ_akt.docx'
            PRICE = PRICESZ
        if record[7]:
            context['kolweb'] = record[7]
            context['sumweb'] = round(int(record[7]) * PRICE[0], 2)
            allsum = allsum + float(context['sumweb'])
        else:
            context['kolweb'] = ' '
            context['sumweb'] = ' '
        if record[9]:
            context['kolos'] = record[9]
            context['sumos'] = round(int(record[9]) * PRICE[1], 2)
            allsum = allsum + float(context['sumos'])
        else:
            context['kolos'] = ' '
            context['sumos'] = ' '
        if record[11]:
            context['kolvst'] = record[11]
            context['sumvst'] = round(int(record[11]) * PRICE[2], 2)
            allsum = allsum + float(context['sumvst'])
        else:
            context['kolvst'] = context['sumvst'] = ' '
        if record[13]:
            context['kolpoe'] = record[13]
            context['sumpoe'] = round(int(record[13]) * PRICE[2], 2)
            allsum = round(allsum + float(context['sumpoe']), 2)
        else:
            context['kolpoe'] = ' '
            context['sumpoe'] = ' '
        context['allsum'] = str(allsum)
        context['allsumbk'] = str(int(allsum))
        context['sumprop'] = num2words(int(allsum), lang='ru')
        context['kop'] = int((allsum * 100) % 100)
        m3 = morph.parse('рубль')[0]
        context['rub'] = m3.make_agree_with_number(int(allsum)).word
        context['filelink'] = filelink
        context['filename'] = name
        doc = DocxTemplate(templ_name1)
        doc.render(context)
        doc.save(filename + ' ' + prilmonth + prilyear + '_приложение.docx')
        doc = DocxTemplate(templ_name2)
        doc.render(context)
        doc.save(filename + ' ' + prilmonth + prilyear + '_акт.docx')

record = []
wb = openpyxl.load_workbook('work.xlsx', data_only=True)
sheet = wb.get_sheet_by_name('sheet1')
num = 2
for row in sheet['A2':'P39']:
    for cellObj in row:
        record.append(cellObj.value)
    make_file(record)
    record = []

import openpyxl
from docxtpl import DocxTemplate
import pymorphy2
from num2words import num2words
import sys
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow
day30 = ['апрель', 'июнь', 'сентябрь', 'ноябрь']
day31 = ['январь', 'март', 'май', 'июль', 'август', 'октябрь', 'декабрь']
PRICEFL = [2874.0, 1724.1, 6896.5]
PRICESZ = [2659.6, 1595.7, 6383.0]
morph = pymorphy2.MorphAnalyzer()


class MyWidget(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('design.ui', self)
        self.make_report.clicked.connect(self.run)

    def run(self):
        self.record = []
        self.prilmonth = self.monthLine.text()
        self.prilyear = self.yearLine.text()
        self.filelink = self.linkLine.text()
        self.name = self.nameLine.text()
        self.m = morph.parse(self.prilmonth)[0]
        wb = openpyxl.load_workbook('work.xlsx', data_only=True)
        sheet = wb.get_sheet_by_name('sheet1')
        num = 2
        for row in sheet['A2':'P39']:
            for cellObj in row:
                self.record.append(cellObj.value)
            self.make_file(self.record)
            self.record = []
        self.label.setText("OK")

    def make_file(self, record):
        # если методист в этот месяце что-то делал
        if record[-1]:
            initial = record[6].split()
            initial = initial[0] + ' ' + initial[1][0] + '. ' + initial[2][0] + '.'
            m2 = morph.parse(record[0])[0]
            # формируем общую часть шаблона
            context = {'dogmonth': m2.inflect({'sing', 'gent'}).word, 'dogyear': record[1], 'prilnum': record[3],
                       'subj': record[4], 'dognum': record[5], 'fio': record[6], 'finit': initial,
                       'prilmonthrp': self.m.inflect({'sing', 'gent'}).word,
                       'prilmonthpp': self.m.inflect({'sing', 'loct'}).word,
                       'prilyear': self.prilyear}
            if self.prilmonth in day30:
                context['lastday'] = '30'
            elif self.prilmonth in day31:
                context['lastday'] = '31'
            else:
                context['lastday'] = '28'
            filename = self.record[6].split()[0]
            allsum = 0
            if self.record[2] == 'ФЛ':
                templ_name1 = 'template.docx'
                templ_name2 = 'template_akt.docx'
                templ_name3 = 'fulltemplate.docx'
                PRICE = PRICEFL
            else:
                templ_name1 = 'templateSZ.docx'
                templ_name2 = 'templateSZ_akt.docx'
                templ_name3 = 'fulltemplateSZ.docx'
                PRICE = PRICESZ
            if record[7]:
                context['kolweb'] = self.record[7]
                context['sumweb'] = round(int(self.record[7]) * PRICE[0], 2)
                allsum = allsum + float(context['sumweb'])
            else:
                context['kolweb'] = ' '
                context['sumweb'] = ' '
            if self.record[9]:
                context['kolos'] = self.record[9]
                context['sumos'] = round(int(self.record[9]) * PRICE[1], 2)
                allsum = allsum + float(context['sumos'])
            else:
                context['kolos'] = ' '
                context['sumos'] = ' '
            if self.record[11]:
                context['kolvst'] = self.record[11]
                context['sumvst'] = round(int(self.record[11]) * PRICE[2], 2)
                allsum = allsum + float(context['sumvst'])
            else:
                context['kolvst'] = context['sumvst'] = ' '
            if self.record[13]:
                context['kolpoe'] = self.record[13]
                context['sumpoe'] = round(int(self.record[13]) * PRICE[2], 2)
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
            context['filelink'] = self.filelink
            context['filename'] = self.name
            doc = DocxTemplate(templ_name1)
            doc.render(context)
            doc.save('parts/' + filename + ' ' + self.prilmonth + self.prilyear + '_приложение.docx')
            doc = DocxTemplate(templ_name2)
            doc.render(context)
            doc.save('parts/' + filename + ' ' + self.prilmonth + self.prilyear + '_акт.docx')
            doc = DocxTemplate(templ_name3)
            doc.render(context)
            doc.save('full/' + filename + ' ' + self.prilmonth + self.prilyear + '.docx')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyWidget()
    ex.show()
    sys.exit(app.exec_())




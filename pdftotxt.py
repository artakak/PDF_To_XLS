#!-*-coding:utf-8-*-
from cStringIO import StringIO
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
import time
import datetime
import re
import xlwt
import pickle
import sys
import shelve
import os
# import PyQt4 QtCore and QtGui modules
from PyQt4.QtCore import *
from PyQt4.QtGui import *
# from PyQt4 import uic
from ui_test import Ui_MainWindow

#( Ui_MainWindow, QMainWindow ) = uic.loadUiType( 'ui_test.ui' )

class MainWindow(QMainWindow):
    """MainWindow inherits QMainWindow"""

    def __init__(self, parent=None):
        QMainWindow.__init__(self, parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.fname = None
        self.page = None
        self.ui.toolButton.clicked.connect(self.fname_pages_dialog)
        self.ui.pushButton.clicked.connect(self.convert)
        self.DB = shelve.open('DB.txt')
        self.DB['Template1'] = ['\d\d-\d\d-\d\d','\n(Tue|Wed|Fri|Thu|Mon|Sat|Sun)','[^\n:]\d\d:\d\d','Voice','\d{7,}',r'\bL\b',r'\b(1|2|3)\b\n',r'\bO-P\b',r'\b00:\d\d:\d\d\b',r'\bH:M:S\b',r'\b\d+,\d+\b']
        self.DB.close()
        self.enum_temp()
        self.ui.comboBox.currentIndexChanged.connect(self.enum_temp)
        self.ui.pushButton_2.clicked.connect(self.save_temp)


    def enum_temp(self):
        template = str(self.ui.comboBox.currentText())
        self.DB = shelve.open('DB.txt')
        self.ui.lineEdit_2.setText(str(self.DB[template][0]))
        self.ui.lineEdit_3.setText(str(self.DB[template][1]))
        self.ui.lineEdit_4.setText(str(self.DB[template][2]))
        self.ui.lineEdit_5.setText(str(self.DB[template][3]))
        self.ui.lineEdit_6.setText(str(self.DB[template][4]))
        self.ui.lineEdit_7.setText(str(self.DB[template][5]))
        self.ui.lineEdit_8.setText(str(self.DB[template][6]))
        self.ui.lineEdit_9.setText(str(self.DB[template][7]))
        self.ui.lineEdit_10.setText(str(self.DB[template][8]))
        self.ui.lineEdit_11.setText(str(self.DB[template][9]))
        self.ui.lineEdit_12.setText(str(self.DB[template][10]))
        self.DB.close()

    def save_temp(self):
        self.DB = shelve.open('DB.txt')
        template = str(self.ui.comboBox.currentText())
        if template != 'Template1':
            self.DB[template] = [str(self.ui.lineEdit_2.text()), str(self.ui.lineEdit_3.text()), str(self.ui.lineEdit_4.text()), str(self.ui.lineEdit_5.text()), str(self.ui.lineEdit_6.text()), str(self.ui.lineEdit_7.text()), str(self.ui.lineEdit_8.text()), str(self.ui.lineEdit_9.text()), str(self.ui.lineEdit_10.text()), str(self.ui.lineEdit_11.text()), str(self.ui.lineEdit_12.text())]
        self.DB.close()

    def fname_pages_dialog(self):
        self.fname = QFileDialog.getOpenFileName(self,'Open PDF','/')
        self.pages = []
        try:
            self.truepage = QInputDialog.getText(self,'Page select','Enter page number')
            for tp in list(self.truepage[0].split(',')):
                self.pages.append(int(tp))
        except: pass
        self.ui.lineEdit.setText(str(self.fname))


    def convert(self):
        self.ui.progressBar.setValue(0)
        if not self.pages:
            pagenums = set()
        else:
            pagenums = set(self.pages)

        output = StringIO()
        manager = PDFResourceManager()
        converter = TextConverter(manager, output, laparams=LAParams())
        interpreter = PDFPageInterpreter(manager, converter)

        infile = file(self.fname, 'rb')
        for page in PDFPage.get_pages(infile, pagenums):
            interpreter.process_page(page)
        infile.close()
        converter.close()
        data = output.getvalue()
        output.close

        #print(data)

        style1 = xlwt.easyxf('font: bold off; align: wrap off, vert centre, horiz center; borders: top thin, bottom thin, left thin, right thin;')
        style1.num_format_str = 'DD-MM-YY'
        style2 = xlwt.easyxf('font: bold off; align: wrap off, vert centre, horiz center; borders: top thin, bottom thin, left thin, right thin;')
        style2.num_format_str = 'HH:MM'
        style4 = xlwt.easyxf('font: bold off; align: wrap off, vert centre, horiz center; borders: top thin, bottom thin, left thin, right thin;')
        style4.num_format_str = 'HH:MM:SS'
        style0 = xlwt.easyxf('font: bold off; align: wrap off, vert centre, horiz left; borders: top thin, bottom thin, left thin, right thin;')
        style3 = xlwt.easyxf('font: bold on; align: wrap off, vert centre, horiz center; borders: top double, bottom double, left double, right double;')
        style5 = xlwt.easyxf('font: bold on; align: wrap off, vert centre, horiz center; borders: top double, bottom double, left double, right double;')
        style5.num_format_str = '[h]:mm:ss;@'

        wb = xlwt.Workbook()
        ws = wb.add_sheet('A Test Sheet')
        ws.write(3,0,'Day',style3)
        ws.write(3,1,'Date',style3)
        ws.write(3,2,'Time',style3)
        ws.write(3,3,'E/Stn',style3)
        ws.write(3,4,'Service Name',style3)
        ws.write(3,5,'Destination',style3)
        ws.write(3,6,'Code',style3)
        ws.write(3,7,'Band',style3)
        ws.write(3,8,'Peak/Off-Peak',style3)
        ws.write(3,9,'Amount',style3)
        ws.write(3,10,'Unit',style3)
        ws.write(3,11,'Cost',style3)
        ws.write(3,12,'Tarif',style3)
        ws.write(3,13,'Cost_2',style3)

        i = 4

        self.ui.progressBar.setValue(10)
        #regDate = '\d\d-\d\d-\d\d'
        regDate = str(self.ui.lineEdit_2.text())
        matchesDate = re.findall(regDate, data)
        #print(matchesDate)
        #print(len(matchesDate))
        for Date in matchesDate:
            #outfp.write(Date+'\n')
            ws.write(i,1,Date,style1)
            i +=1

        i = 4

        #regDay = '\n(Tue|Wed|Fri|Thu|Mon|Sat|Sun)'
        regDay = str(self.ui.lineEdit_3.text())
        matchesDay = re.findall(regDay, data)
        for Day in matchesDay:
            #outfp.write(Day+'\n')
            ws.write(i,0,Day,style0)
            i +=1

        i = 4
        k = 0

        self.ui.progressBar.setValue(20)
        #regTime = '[^\n:]\d\d:\d\d'
        regTime = str(self.ui.lineEdit_4.text())
        matchesTime = re.findall(regTime, data)
        #print(matchesTime)
        #print(len(matchesTime))
        while k < len(matchesDate):
            #outfp.write(Time+'\n')
            ws.write(i,2,matchesTime[k],style2)
            k+=1
            i+=1

        i = 4
        k = 0

        #regService = 'Voice'
        regService = str(self.ui.lineEdit_5.text())
        matchesService = re.findall(regService, data)
        #print(matchesService)
        #print(len(matchesService))
        while k < len(matchesDate):
            #outfp.write(Day+'\n')
            ws.write(i,4,matchesService[k],style0)
            i+=1
            k+=1

        i = 4
        k = 0

        self.ui.progressBar.setValue(30)
        trueDest = []
        truedest2 = []
        regDest = str(self.ui.lineEdit_6.text())
        #regDest = '\d{7,}'
        matchesDest = re.findall(regDest, data)
        ws.write(0,0,'Mobile No',style3)
        ws.write(0,1,matchesDest[0],style3)
        #print(matchesDest)
        #print(len(matchesDest))
        for Dest in matchesDest:
            if Dest != matchesDest[0]:
                #outfp.write(Day+'\n')
                trueDest.append(Dest)
        while k < len(matchesDate):
            ws.write(i,5,trueDest[k],style0)
            truedest2.append(trueDest[k][0:4])
            i+=1
            k+=1

        #print(len(truedest2))
        #print(truedest2)
        i = 4
        k = 0

        self.ui.progressBar.setValue(40)
        regCode = str(self.ui.lineEdit_7.text())
        #regCode = r'\bL\b'
        matchesCode = re.findall(regCode, data)
        #print(matchesCode)
        #print(len(matchesCode))
        while k < len(matchesDate):
            #outfp.write(Day+'\n')
            ws.write(i,6,matchesCode[k],style0)
            i+=1
            k+=1

        i = 4
        k = 0

        self.ui.progressBar.setValue(50)
        regBand = str(self.ui.lineEdit_8.text())
        #regBand = r'\b(1|2|3)\b\n'
        matchesBand = re.findall(regBand, data)
        #print(matchesBand)
        #print(len(matchesBand))
        while k < len(matchesDate):
            #outfp.write(Day+'\n')
            ws.write(i,7,matchesBand[k],style0)
            i+=1
            k+=1

        i = 4
        k = 0

        self.ui.progressBar.setValue(60)
        regPeak = str(self.ui.lineEdit_9.text())
        #regPeak = r'\bO-P\b'
        matchesPeak = re.findall(regPeak, data)
        #print(matchesPeak)
        #print(len(matchesPeak))
        while k < len(matchesDate):
            #outfp.write(Day+'\n')
            ws.write(i,8,matchesPeak[k],style0)
            i+=1
            k+=1

        i = 4
        k = 0

        self.ui.progressBar.setValue(70)
        regAmount = str(self.ui.lineEdit_10.text())
        #regAmount = r'\b00:\d\d:\d\d\b'
        matchesAmount = re.findall(regAmount, data)
        #print(matchesAmount)
        #print(len(matchesAmount))
        while k < len(matchesDate):
            #outfp.write(Day+'\n')
            hms = matchesAmount[k].split(':')
            ws.write(i,9,datetime.time(int(hms[0]),int(hms[1]),int(hms[2])),style4)
            i+=1
            k+=1

        i = 4

        regUnit = str(self.ui.lineEdit_11.text())
        #regUnit = r'\bH:M:S\b'
        matchesUnit = re.findall(regUnit, data)
        #print(matchesUnit)
        #print(len(matchesUnit))
        for Unit in matchesUnit:
            #outfp.write(Day+'\n')
            ws.write(i,10,Unit,style0)
            i+=1

        i = 4
        k = 0

        self.ui.progressBar.setValue(80)
        regCost = str(self.ui.lineEdit_12.text())
        #regCost = r'\b\d+,\d+\b'
        matchesCost = re.findall(regCost, data)
        #print(matchesCost)
        #print(len(matchesCost))
        while k < len(matchesDate):
            #outfp.write(Day+'\n')
            ws.write(i,11,float(matchesCost[k].replace(',','.')),style0)
            i+=1
            k+=1

        i = 4
        k = 0

        self.ui.progressBar.setValue(90)
        tarif = {}
        try:
            tariff = file('tarif.txt', 'r')
            tarif = pickle.load(tariff)
            tariff.close()
            #print(tarif)
        except: pass
        while k < len(matchesDate):
            if tarif.has_key(truedest2[k]):
                #outfp.write(Day+'\n')
                ws.write(i,12,tarif[truedest2[k]],style0)
            else:
                tarifinput = QInputDialog.getText(self,'Tarif','Please enter tarif for '+truedest2[k]+':')
                tarif[truedest2[k]] = float(tarifinput[0])
                ws.write(i,12,tarif[truedest2[k]],style0)

            ws.write(i,13,xlwt.Formula('(HOUR(J'+str(i+1)+')*60+MINUTE(J'+str(i+1)+')+SECOND(J'+str(i+1)+')/60)*M'+str(i+1)+''),style0)
            i+=1
            k+=1


        tariff = file('tarif.txt', 'w')
        pickle.dump(tarif, tariff)
        tariff.close()

        ws.write(len(matchesDate)+6,0,'Total:',style3)
        ws.write(len(matchesDate)+6,9,xlwt.Formula('SUM(J5:J'+str(len(matchesDate)+4)+')'),style5)
        ws.write(len(matchesDate)+6,11,xlwt.Formula('SUM(L5:L'+str(len(matchesDate)+4)+')'),style3)
        ws.write(len(matchesDate)+6,13,xlwt.Formula('SUM(N5:N'+str(len(matchesDate)+4)+')'),style3)


        #outfp.close
        wb.save(self.fname+'.xls')
        self.ui.progressBar.setValue(100)
        if self.ui.checkBox.isChecked():
            os.system(''+str(self.fname)+'.xls')
        return data

    def __del__(self):
        self.ui = None


#-----------------------------------------------------#
if __name__ == '__main__':
    # create application
    app = QApplication(sys.argv)
    app.setApplicationName('PDF_To_XLS')

    # create widget
    w = MainWindow()
    w.setWindowTitle('PDF_To_XLS')
    w.show()

    # connection
    QObject.connect(app, SIGNAL('lastWindowClosed()'), app, SLOT('quit()'))

    # execute application
    sys.exit(app.exec_())
import sys
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QApplication, QDialog
from PyQt5.uic import loadUi
from datetime import datetime
import pyodbc;
import xlwt 
from xlwt import Workbook
class MainPage(QDialog):
    def __init__(self):
        super(MainPage, self).__init__()
        loadUi('main.ui',self)
        ##
        self.pushButton.clicked.connect(self.retrieve)
    def retrieve(self):
        show = self.dateTimeEdit.dateTime()
        show1 = self.dateTimeEdit_2.dateTime()
        show_string = show.toString(self.dateTimeEdit.displayFormat())
        show_string1 = show1.toString(self.dateTimeEdit_2.displayFormat())
        print(show_string.replace("/", "-"))
        lim1 = show_string.replace("/", "-")
        lim2 = show_string1.replace("/", "-")
        print("Por favor espere mientras la información es procesada")
        for i in range(0,3):
            print("...")
        ##CONEXION A FLOATTABLE
        cnxn = pyodbc.connect(driver='{SQL Server}', host='DELL2-PC\FTVIEWX64TAGDB', database='CEE_Report_New',
                              trusted_connection='yes')
        cursor = cnxn.cursor()
        cursor.execute("SELECT * FROM CEE_Report_New.dbo.FloatTable WHERE DateAndTime BETWEEN '%s' AND '%s'" % (lim1,lim2))
        rows_info = cursor.fetchall()
        print(rows_info[0][0])
        cnxn.close()
        ##CONEXION A TAGTABLE
        cnxn_1 = pyodbc.connect(driver='{SQL Server}', host='DELL2-PC\FTVIEWX64TAGDB', database='CEE_Report_New',
                              trusted_connection='yes')
        cursor_1 = cnxn_1.cursor()
        cursor_1.execute("SELECT * FROM CEE_Report_New.dbo.TagTable")
        rows_tag = cursor_1.fetchall()
        cnxn_1.close()
        print("Conexión a la base de datos establecida")
        ## Primera lista
        tags = ['DateAndTime','TT_SPD3001_201.Val', 'TT_SPD3001_202.Val',
                             'TT_SPD3001_203.Val', 'TT_SPD3001_204.Val',
                             'TT_CS3001_205.Val', 'TT_VFBD3001_206.Val',
                             'TT_AH3001_207.Val', 'TT_AH3002_208.Val',
                             'TT_AH3003_209.Val', 'TT_VFBD3001_210.Val',
                             'TT_H3011_211.Val', 'TT_H3011_212.Val',
                             'TT_H3011_213.Val', 'TT_H3011_214.Val',
                             'TT_SCC3005_215.Val', 'TT_SCC3005_216.Val',
                             'TT_VFBD3001_217.Val', 'TT_SPD3001_218.Val',
                             'TT_SPD3001_219.Val', 'TT_SPD3001_220.Val',
                             'TT_F3001_221.Val', 'PT_SPD3001_101.Val',
                             'DPT_F3001_102.Val', 'DPT_SCC3005_103.Val',
                             'DPT_VFBD3001_104.Val', 'PT_VFBD3001_105.Val',
                             'DPT_CS3001_106.Val', 'DPT_SCC3002A_107.Val',
                             'DPT_SCC3002B_108.Val', 'PT_B3001_109.Val',
                             'PT_B3001_110.Val', 'PT_SPD3001_111.Val',
                             'PT_SPD3001_112.Val', 'PT_SPD3001_113.Val',
                             'PT_VFBD3001_114.Val', 'FPT_SPD3001_101.Val',
                             'PT_LP3001_102.Val', 'PT_SCC3002A_103.Val',
                             'PT_SCC3002B_104.Val', 'FFT_SPD3001_301.Val',
                             'GFT_HAG3001_302.Val', 'LT_SCT3002A_101.Val',
                             'LT_SCT3002B_102.Val', 'FCV_AH3001_101.CVEU',
                             'FCV_AH3002_102.CVEU', 'FCV_AH3003_103.CVEU',
                             'CV_VFBD3001_104.CVEU', 'CV_VFBD3001_105.CVEU',
                             'CV_CS3001_106.CVEU', 'CV_SCC3002A_107.CVEU',
                             'CV_SCC3002B_108.CVEU', 'DRIVE_7:I.Data[1]',
                             'DRIVE_2:I.Data[1]', 'DRIVE_6:I.Data[1]',
                             'DRIVE_5:I.Data[1]', 'DRIVE_4:I.Data[1]',
                             'DRIVE_3:I.Data[1]']
        # Workbook is created 
        wb = Workbook()  
        # add_sheet is used to create sheet.
        sheet1 = wb.add_sheet('CEE_Report',cell_overwrite_ok = True)
        style = xlwt.XFStyle()
        style1 = xlwt.XFStyle()
        # font
        font = xlwt.Font()
        font1 = xlwt.Font()
        font.bold = True
        font1.italic = True
        style1.font1 = font1
        style.font = font
        #Write the tag name in the table
        for i in range(0,len(tags)):
            sheet1.write(0,i,tags[i],style) ## Se escriben los tags en las columnas de la primera fila
        ## IMPRIME LOS TIEMPOS EN LA PRIMERA COLUMNA
        hora_list = []
        for i in rows_info:
            var = str(i[2])
            if var == "0":
                hora_list.append(str(i[0]))
        for i in range(0,len(hora_list)):
            hora_red = hora_list[i][0:19] ## Se eliminan los ceros de más en la hora;
            sheet1.write(i+1,0,hora_red) ## Se escriben los tiempos en las filas de la primera columna            
        ## IMPRIME LOS VALORES DE CADA TAG PARA CADA HORA
        o = 0
        l = 57
        for x  in range(1, len(hora_list)+1):
            z = 0
            for y in range(o,l):
                sheet1.write(x,z+1, "%.3f" % float(rows_info[y][3]), style1)## Se escriben los datos fila por fila;
                z = z+1
            o = o+57
            l = l+57
        wb.save('CEE_Report.xls') ## Se guarda el documento;
        print("El documento ha sido creado, por favor revisar la carpeta raíz")
        print("Recuerde eliminar o mover el archivo creado para futuros reportes")
app = QApplication(sys.argv)
widget = MainPage()
widget.show()
sys.exit(app.exec_())

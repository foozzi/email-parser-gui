import sys
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog, QMessageBox
from bs4 import BeautifulSoup
from requests.exceptions import MissingSchema
import re
import requests
import main
import xlsxwriter

class MainWindow(QtWidgets.QMainWindow):
    mails_arr = []    

    
    def __init__(self):
        super(MainWindow, self).__init__()
        self.ui = main.Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.pushButton_3.setEnabled(False)
        self.ui.pushButton_2.setEnabled(False)
        self.ui.pushButton.clicked.connect(self.open_sites)
        self.ui.pushButton_2.clicked.connect(self.search_mails)
        self.ui.pushButton_3.clicked.connect(self.save_excel)
        self.ui.progressBar.hide()
        self.ui.label.hide()
        
    def open_sites(self):        
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","Text doc (*.txt)", options=options)
        # protect xD
        max_sites = 1000
        incr_sites = 0
        if fileName:
            with open(fileName) as f:
                sites = f.readlines()
            sites_list = ''
            for site in sites:
                if incr_sites >= max_sites:
                    break
                sites_list += site.strip() + '\n'
                incr_sites += 1                
            self.ui.progressBar.show()
            self.ui.label.show()
            self.ui.plainTextEdit.setPlainText(sites_list)
            self.ui.pushButton_2.setEnabled(True)
            
    def search_mails(self):     
        
        self.ui.plainTextEdit_2.setPlainText('')
        sites = self.ui.plainTextEdit.toPlainText().strip()
        if sites:
            site_incr = 0
            for site in sites.split('\n'):                       
                try:
                    request = requests.get(site)
                except MissingSchema as e: 
                    self.ui.plainTextEdit.setPlainText('')
                    self.show_alert('Some url is not valid! Check you .txt file')
                    return
                
                site_incr += 1
                progress = (100 / len(sites.split('\n'))) * site_incr
                self.ui.progressBar.setValue(progress)

                if request.status_code != 200:      
                    continue
                html = request.text
                links = self.find_links(html)
                tmp_arr = []
                for link in links:
                    if link.string != None:
                        if link.string.find('@') >= 0:
                            self.ui.plainTextEdit_2.appendPlainText(link.string.strip())
                            tmp_arr.append(link.string)
                            #self.current_mails += link.string.strip() + '\n'
                self.mails_arr.append({site:tmp_arr})
            self.ui.pushButton_3.setEnabled(True)                        
            
    def find_links(self, html):
        soup = BeautifulSoup(html, 'html.parser')
        return soup.findAll('a')
    
    def save_excel(self):
        workbook = xlsxwriter.Workbook('demo.xlsx')
        worksheet = workbook.add_worksheet()
        worksheet.set_column('A:A', 20)
        bold = workbook.add_format({'bold': True})
        mails = self.mails_arr
        site_incr = 1
        if len(mails) > 0:
            mail_incr = 1
            site_incr = 0
            col = 0
            row = 0
            for mail in mails:                
                for k in mail:
                    if len(mail[k]) < 1:
                        continue
                    worksheet.write('A' + str(site_incr), k)
                    
                    for m in mail[k]:                        
                        worksheet.write(row, col,     k)
                        worksheet.write(row, col + 1, m)
                        row +=1
                        #worksheet.write('A' + str(mail_incr), m)
            workbook.close()
            
    def show_alert(self, text):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText(text)
        msg.exec_()

app = QtWidgets.QApplication(sys.argv)

my_mainWindow = MainWindow()
my_mainWindow.show()

sys.exit(app.exec_())

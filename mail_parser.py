import sys
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog, QMessageBox
from bs4 import BeautifulSoup
from requests.exceptions import MissingSchema
from urllib.parse import urlparse
from email.utils import parseaddr
import re
from multiprocessing import Queue
import requests
import main
import xlsxwriter

class MainWindow(QtWidgets.QMainWindow):
    mails_arr = []   
    tmp_arr = []

    
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
                site_incr += 1
                progress = (100 / len(sites.split('\n'))) * site_incr
                self.ui.progressBar.setValue(progress)
                html = self.request_site(site)
                if not html:
                    continue
                
                links = self.find_links(html)
                self.tmp_arr = []
                for link in links:
                    if link.string != None:
                        self.detect_email(link)
                        
                        if link['href'] == None:
                            continue
                        
                        o = urlparse(link['href'])
                        
                        m = re.search('/' + o.netloc + '/', link['href'])
                        if m != None and o.scheme != '':
                            print(link['href'])
                            html = self.request_site(link['href'])
                            links = self.find_links(html)
                            for link in links:
                                self.detect_email(link)                                
                            
                self.mails_arr.append({site:self.tmp_arr})
            self.ui.pushButton_3.setEnabled(True)   
            
    def detect_email(self, string):
        mail = parseaddr(string.string)                            
        if mail[1].find('@') >= 0:              
            if mail[1].lower() in self.tmp_arr:
                return False
            self.ui.plainTextEdit_2.appendPlainText(mail[1].strip().lower())
            self.tmp_arr.append(mail[1].lower())
            
    def request_site(self, url):
        try:
            request = requests.get(url)
        except MissingSchema as e: 
            self.ui.plainTextEdit.setPlainText('')
            self.show_alert('Some url is not valid! Check you .txt file')
            return
        
        if request.status_code != 200:      
            return False
        else:
            return request.text
            
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

from selenium import webdriver
from time import time, sleep
import pandas as pd
from openpyxl import Workbook
import xlsxwriter
import time


class Currency:
    def __init__(self, amount):
        self.url="https://anlikaltinfiyatlari.com/"
        self.browser=webdriver.Chrome(executable_path="C:\seleniumbrowserdriverSave\chromedriver.exe")
        self.amount=amount

    def currencyOns(self):
       while True:
        self.browser.get(self.url)
        time.sleep(3)
        day=self.browser.find_element_by_xpath("//*[@id='content']/div/h3").text
        onsAltın=self.browser.find_element_by_xpath("//*[@id='spot']/ul/li[2]").text
        gramAltın=self.browser.find_element_by_xpath("//*[@id='spot']/ul/li[3]/a").text
        dolar=self.browser.find_element_by_xpath("//*[@id='spot']/ul/li[4]").text
        euro=self.browser.find_element_by_xpath("//*[@id='spot']/ul/li[5]").text
        euroDdolar=self.browser.find_element_by_xpath("//*[@id='spot']/ul/li[6]").text
        sterlin=self.browser.find_element_by_xpath("//*[@id='sidebar']/div[1]/ul/li[4]").text
        frank=self.browser.find_element_by_xpath("//*[@id='sidebar']/div[1]/ul/li[5]").text
        riyal=self.browser.find_element_by_xpath("//*[@id='sidebar']/div[1]/ul/li[6]").text
        bulgarLevası=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/div/div/ul/li[7]").text
        manat=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/div/div/ul/li[8]").text
        japonYeni=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/div/div/ul/li[9]").text
        rusRublesi=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/div/div/ul/li[10]").text
        avusturyaDoları=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/div/div/ul/li[11]").text
        ceyrekAltın=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/table/tbody/tr[4]").text
        yarımAltın=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/table/tbody/tr[5]").text
        tamAltın=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/table/tbody/tr[6]").text
        cumhuriyetAltın=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/table/tbody/tr[7]").text
        gümüs=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/table/tbody/tr[15]").text
        gümüsOns=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/table/tbody/tr[16]").text

        time.sleep(60)
        print(f"{day} \n,{onsAltın} \n,{gramAltın} \n,{dolar} \n,{euro} \n,{euroDdolar} \n,{sterlin} \n,"
              f" {frank} \n,{riyal},{bulgarLevası} \n,{manat} \n,{japonYeni} \n,{rusRublesi} \n,"
              f" {avusturyaDoları} \n,{ceyrekAltın} \n,{yarımAltın} \n,{tamAltın} \n,{cumhuriyetAltın} \n,"
              f" {gümüs} \n,{gümüsOns} \n")

        # Yukarıdaki aşamalarda öncelikle internet üzerinde adrese ulaşıp o kısımdaki döviz bilgilerini alıyoruz.
        # Sonrasında 60 saniyede bir güncel veriler karşımıza gelecek şekilde çalışmaktadır.
        # Çıkmış olan verileri Excel'e yazmamız gerekmektedir.

        #with workbook() as wb:
        wb=Workbook()
        wb['Sheet'].title="Report"
        sh1=wb.active
        sh1['A1'].value="Day "
        sh1['C1'].value=self.browser.find_element_by_xpath("//*[@id='content']/div/h3").text
        sh1['A2'].value="Ons Altın "
        sh1['C2'].value=self.browser.find_element_by_xpath("//*[@id='spot']/ul/li[2]").text
        sh1['A3'].value="Gram Altın "
        sh1['C3'].value=self.browser.find_element_by_xpath("//*[@id='spot']/ul/li[3]/a").text
        sh1['A4'].value="Dolar "
        sh1['C4'].value=self.browser.find_element_by_xpath("//*[@id='spot']/ul/li[4]").text
        sh1['A5'].value="Euro "
        sh1['C5'].value=self.browser.find_element_by_xpath("//*[@id='spot']/ul/li[5]").text
        sh1['A6'].value="Euro / Dolar "
        sh1['C6'].value=self.browser.find_element_by_xpath("//*[@id='spot']/ul/li[6]").text
        sh1['A7'].value="Sterlin "
        sh1['C7'].value=self.browser.find_element_by_xpath("//*[@id='sidebar']/div[1]/ul/li[4]").text
        sh1['A8'].value="Frank "
        sh1['C8'].value=self.browser.find_element_by_xpath("//*[@id='sidebar']/div[1]/ul/li[5]").text
        sh1['A9'].value="Riyal "
        sh1['C9'].value=self.browser.find_element_by_xpath("//*[@id='sidebar']/div[1]/ul/li[6]").text
        sh1['A10'].value="Bulgar Levası "
        sh1['C10'].value=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/div/div/ul/li[7]").text
        sh1['A11'].value="Manat "
        sh1['C11'].value=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/div/div/ul/li[8]").text
        sh1['A12'].value="Japon Yeni "
        sh1['C12'].value=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/div/div/ul/li[9]").text
        sh1['A13'].value="Rus Rublesi "
        sh1['C13'].value=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/div/div/ul/li[10]").text
        sh1['A14'].value="Avusturya Doları "
        sh1['C14'].value=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/div/div/ul/li[11]").text
        sh1['A4'].value="Ceyrek Altın "
        sh1['C4'].value=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/table/tbody/tr[4]").text
        sh1['A5'].value="Yarım Altın "
        sh1['C5'].value=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/table/tbody/tr[5]").text
        sh1['A6'].value="Tam Altın "
        sh1['C6'].value=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/table/tbody/tr[6]").text
        sh1['A1'].value="Cumhuriyer Altın "
        sh1['C1'].value=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/table/tbody/tr[7]").text
        sh1['A2'].value="Gümüş "
        sh1['C2'].value=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/table/tbody/tr[15]").text
        sh1['A3'].value="Gümüş Ons "
        sh1['C3'].value=self.browser.find_element_by_xpath("//*[@id='content']/div/div[2]/table/tbody/tr[16]").text

        wb.save(r"C:\\Users\\Murat Can Özen\\Desktop\\dovizIslemEx.xlsx")
        # Bu kısımda internet üzerinden çekmiş olduğumuz verileri Excel üzerine yazmaktayız.


currency=Currency("Döviz Değişim")
currency.currencyOns()

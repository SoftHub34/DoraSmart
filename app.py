from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from task_base import TaskBase
import base64
from twocaptcha import TwoCaptcha
from imagetyperzapi3.imagetyperzapi import ImageTyperzAPI
import time


class ChaptcaSolver:
    def __init__(self, resolve_min_len, resolve_max_len):
        self.resolve_min_len = resolve_min_len
        self.resolve_max_len = resolve_max_len
        twochaptca_key = '133426027087732629770816440575179997367'
        typerz_key = '167384896567667411170845136371732868948'    
        twochaptca_key = hex(int(twochaptca_key))[2:]
        typerz_key = hex(int(typerz_key))[2:].upper()
        self.two_solver = TwoCaptcha(twochaptca_key)
        self.typerz_solver = ImageTyperzAPI(typerz_key)


    def solve(self, img_path):
        solver = self.two_solver
        try:
            result = solver.normal(img_path)
            if result != None and len(result['code']) >= self.resolve_min_len and len(result['code']) <= self.resolve_max_len:
                return result['code']
            else:
                solver = self.typerz_solver
                result = solver.solve_captcha(img_path)
                if result != None and len(result) >= self.resolve_min_len and len(result) <= self.resolve_max_len:
                    return result
                else:
                    return ''
        except:
            return ''







class Task(TaskBase):
    def __init__(self):
        super().__init__()
        self.selenium = Selenium()
        self.excel = Files()
        self.chaptca=ChaptcaSolver(5,5)
        
     
    
    
   # def tarih(self):
    #    self.selenium.open_available_browser("https://dev.smartportal.com.tr/smart-robot-zamanlayici")
     #   self.start=self.selenium.get_value(locator="xpath://*[@id='mat-input-2']")
      #  self.end=self.selenium.get_value(locator="xpath://*[@id='mat-input-2']")#siteden alınacak
        
            
    def Giris_bilgileri_alma(self):
        self.selenium.open_available_browser("https://dev.smartportal.com.tr/isyeri/1a6d90bd-499f-45fd-9e62-14189bed53da")
        self.selenium.input_text(locator="xpath://*[@autocomplete='email']",text="system@smartportal.com")
        self.selenium.input_text(locator="xpath://*[@autocomplete='current-password']",text="smartportal")
        self.selenium.click_button(locator="xpath://*[@class='mat-focus-indicator w-100 mat-raised-button mat-button-base mat-primary']")
        time.sleep(5)
        self.selenium.click_element(locator="xpath://*[@routerlink='/isyeri']")
        time.sleep(1)
        self.selenium.click_element(locator="xpath://*[@id='main-sidenav']/div/app-isyeri-list-page/div/app-isyeri-list-table/div/table/tbody/tr[5]/td[7]/button[1]/span[1]/mat-icon")
        time.sleep(2)
        
        self.kullanici_ad=self.selenium.get_webelement(locator="xpath://*[@class='mat-cell cdk-cell cdk-column-kullaniciAdi mat-column-kullaniciAdi ng-star-inserted']")
        self.kullanici_ad=self.kullanici_ad.get_attribute('innerHTML')
        self.kullanici_ad= self.kullanici_ad[1:len(self.kullanici_ad)-1]
        print("kullanici_adi:",self.kullanici_ad)
        
        self.isyeri_kodu=self.selenium.get_webelement(locator="xpath://*[@class='mat-cell cdk-cell cdk-column-isyeriKodu mat-column-isyeriKodu ng-star-inserted']")
        self.isyeri_kodu=self.isyeri_kodu.get_attribute('innerHTML')
        self.isyeri_kodu= self.isyeri_kodu[1:len(self.isyeri_kodu)-1]
        print("İsyeri_kodu",self.isyeri_kodu)
        
        
        self.sistem_sifre=self.selenium.get_webelement(locator="xpath://*[@class='mat-cell cdk-cell cdk-column-sistemSifresi mat-column-sistemSifresi ng-star-inserted']")
        self.sistem_sifre=self.sistem_sifre.get_attribute('innerHTML')
        self.sistem_sifre= self.sistem_sifre[1:len(self.sistem_sifre)-1]

        print("sistem şifre",self.sistem_sifre)
        
        self.isyeri_sifre=self.selenium.get_webelement(locator="xpath://*[@class='mat-cell cdk-cell cdk-column-isyeriSifresi mat-column-isyeriSifresi ng-star-inserted']")
        self.isyeri_sifre=self.isyeri_sifre.get_attribute('innerHTML')
        self.isyeri_sifre= self.isyeri_sifre[1:len(self.isyeri_sifre)-1]

        print("işyeri şifre",self.isyeri_sifre)
        
        self.ünvan=self.selenium.get_webelement(locator="xpath://*[@id='main-sidenav']/div/app-isyeri-view/div/div/app-view-item-text[3]/div/div[2]")
        self.ünvan=self.ünvan.get_attribute('innerHTML')
        self.ünvan= self.ünvan[1:len(self.ünvan)-1]
        print("ünvan",self.ünvan)
        
        self.sicil_no=self.selenium.get_webelement(locator="xpath://*[@id='main-sidenav']/div/app-isyeri-view/div/div/app-view-item-text[1]/div/div[2]")
        self.sicil_no=self.sicil_no.get_attribute('innerHTML')
        print(self.sicil_no)
        print(type(self.sicil_no))
        #self.sicil_no=int(self.sicil_no)
        self.ayrac=self.sicil_no[23:26]
        print(self.ayrac)     
        print(type(self.ayrac))   
        self.selenium.close_browser()
        
        
    def Giris_bilgileri_yazma(self):    
        self.selenium.open_available_browser("https://uyg.sgk.gov.tr/IsverenSistemi")
        self.selenium.input_text(locator="xpath://*[@id='kullaniciIlkKontrollerGiris_username']",text=self.kullanici_ad)
        self.selenium.input_text(locator="xpath://*[@name='isyeri_kod']",text=self.isyeri_kodu)
        self.selenium.input_text(locator="xpath://*[@ name='password']",text=self.sistem_sifre)
        self.selenium.input_text(locator="xpath://*[@name='isyeri_sifre']",text=self.isyeri_sifre)
        
        img=self.selenium.get_webelement(locator="xpath://*[@id='guvenlik_kod']")
        url=img.get_attribute("src")
        self.chaptca_text = self.chaptca.solve(url) #URL KONTROL
        
        print(url)
        print(self.chaptca_text)
        
        #yer=self.selenium.get_webelement(locator="xpath://*[@id='kullaniciIlkKontrollerGiris_isyeri_guvenlik']")
        #yer=yer.set_attr("value",self.chaptca_text)
        time.sleep(10)
        #self.selenium.click_button("Giriş Yap")
        #self.selenium.click_element(locator="xpath://*[@id='menuForm:anaMenu1subMenu2subItem1']")

        ##self.selenium.click_element(locator="xpath://*[@id='menuForm:anaMenu1subMenu2subItem1']/span")
     ##siteden veri çekilecek   
   
   
    
    def sicil(self):
        self.is_yeri_calisan_sayisi=0
        self.üst_dosya_calisan_sayisi=0
        self.alt_isveren_calisan_sayisi=0
        self.mükerrer_bildirim=0
        self.tekillestirilmis_toplam_sayi=0
        self.selenium.switch_window('İşveren Sistemi',1000)
       
        self.selenium.maximize_browser_window()
        time.sleep(5)
        
        self.selenium.click_element(locator="xpath://*[@id='menuForm:panelMenuId']/h3[2]/span")
        self.selenium.wait_until_element_is_visible(locator="xpath://*[@id='menuForm:anaMenu1subMenu2']/span[1]")
        self.selenium.click_element(locator="xpath://*[@id='menuForm:anaMenu1subMenu2']/span[1]")
        self.selenium.wait_until_element_is_visible(locator="xpath://*[@id='menuForm:anaMenu1subMenu2subItem1']/span")
        
        self.selenium.click_element(locator="xpath://*[@id='menuForm:anaMenu1subMenu2subItem1']/span")
        time.sleep(20)
        #self.selenium.click_button(locator="xpath://*[@id='ortalamaCalisan_sent_donem_yil_ay_index']")
        #self.selenium.click_element(locator="xpath://*[@id='ortalamaCalisan_sent_donem_yil_ay_index']/option[4]")
##tarih
        #self.selenium.click_element(locator="xpath://*[@id='ortalamaCalisan_sent_donem_yil_ay_index']")
        #self.selenium.click_element(locator="xpath://*[@value='1']")        
        #self.selenium.click_button("Listele")
#IsVerenUygulamaPopup
        time.sleep(25)

                ##BAŞKA DOSYADA XPATH KULLANMADAN
        if self.ayrac== "000":
            self.ünvan2=self.selenium.get_text(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[2]/td[1]/p")
            if self.ünvan ==self.ünvan2.get_attribute("innerHTML") :
                self.ünvandeğer=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[i]/td[3]/p")
                self.is_yeri_calisan_sayisi += self.ünvandeğer.get_attribute("innerHTML")
                
            self.üst_dosya_calisan_sayisi += 0
            print(self.üst_dosya_calisan_sayisi)

            
            self.AltİsverenCalisanSayısı2 = self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[i]/td[1]/p")
            if "Alt işveren çalışan sayısı"  == self.AltİsverenCalisanSayısı2.get_attribute("innerHTML"): 
                self.alt_isveren_calisan_sayisi_değer=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[i]/td[1]/p")
                self.alt_isveren_calisan_sayisi+= self.alt_isveren_calisan_sayisi_değer.get_attribute("innerHTML")
    
            #tekrar eden
            self.mükerrer_bildirim+=0

            
            self.tekilleştirilmiş_toplam_sayı2 = self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[i]/td[1]/p")
            if "Tekilleştirilmiş Toplam Sayı"  == self.tekilleştirilmiş_toplam_sayı2.get_attribute("innerHTML"): 
                self.tekilleştirilmiş_değer=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[i]/td[1]/p")
                self.tekillestirilmis_toplam_sayi+= self.tekilleştirilmiş_değer.get_attribute("innerHTML")
            
                    
            if self.ayrac != "000":
            #000 olmayan    
                self.ünvan2=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[i]/td[1]/p")
                if self.ünvan ==self.ünvan2.get_attribute("innerHTML"):
                    self.ünvandeğer=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[i]/td[3]/p")
                    self.is_yeri_calisan_sayisi += self.ünvandeğer.get_attribute("innerHTML")
                
                self.üst_dosya_calisan_sayisi2 = self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[i]/td[1]/p")
                if "Üst Dosya Çalışanı" == self.üst_dosya_calisan_sayisi2.get_attribute("innerHTML"):
                    self.üst_dosya_değer = self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[2]/td[3]/p")
                    self.üst_dosya_calisan_sayisi += self.üst_dosya_değer.get_attribute("innerHTML")

            
                self.alt_isveren_calisan_sayisi+= 0
    
            
                self.tekilleştirilmiş_toplam_sayı2 = self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[i]/td[1]/p") 
                if "Tekilleştirilmiş Toplam Sayı" == self.tekilleştirilmiş_toplam_sayı2.get_attribute("innerHTML"): 
                    self.tekilleştirilmiş_değer=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[i]/td[3]/p")
                    self.tekillestirilmis_toplam_sayi +=self.tekilleştirilmiş_değer.get_attribute("innerHTML")
                
                self.mükerrer_bildirim2 = self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[i]/td[1]/p")
                if "Mükerrer bildirim"  == self.mükerrer_bildirim2.get_attribute("innerHTML"): 
                    self.tekilleştirilmiş_değer=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[i]/td[1]/p")
                    self.tekillestirilmis_toplam_sayi+= self.tekilleştirilmiş_değer.get_attribute("innerHTML")
                
            else:
                i+=1
            
        
        
    def excel(self):
        self.excel.create_workbook()
        self.excel.set_cell_value(1, "A", self.is_yeri_calisan_sayisi )
        self.excel.set_cell_value(1, "B", self.üst_dosya_calisan_sayisi)
        self.excel.set_cell_value(1, "C", self.alt_isveren_calisan_sayisi)
        self.excel.set_cell_value(1, "D", self.mükerrer_bildirim )
        self.excel.set_cell_value(1, "E", self.tekillestirilmis_toplam_sayi)

        self.excel.save_workbook("calışan_sayısı.xlsx")
        
    def run(self):
        Task.Giris_bilgileri_alma(self)
        Task.Giris_bilgileri_yazma(self)
        Task.sicil(self)
        Task.excel(self)

  

        
    


        
                
if __name__ == "__main__":
    #Task.tarih(self)
    Task().run()
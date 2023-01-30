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
        twochaptca_key = '338706026024060347165338827078421340543'
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
        
     
    
        #self.selenium.open_available_browser("https://dev.smartportal.com.tr/smart-robot-zamanlayici")#veri alıncak sayfa
        
       # self.sirket=self.selenium.get_webelement(locator="xpath://*[@class='mat-cell cdk-cell cdk-column-kullaniciAdi mat-column-kullaniciAdi ng-star-inserted']").get_attribute('innerHTML').strip()
        
        #self.start_date=self.selenium.get_webelement(locator="xpath://*[@class='mat-cell cdk-cell cdk-column-kullaniciAdi mat-column-kullaniciAdi ng-star-inserted']").get_attribute('innerHTML').strip()
        
        #self.end_date=self.selenium.get_webelement(locator="xpath://*[@class='mat-cell cdk-cell cdk-column-kullaniciAdi mat-column-kullaniciAdi ng-star-inserted']").get_attribute('innerHTML').strip()
        self.start_date="2022"
        self.end_date="2016"
        self.sirket="DORA SMART İNSAN KAYNAKLARI LTD.ŞTİ."
        self.excel.create_workbook("asgari.xlsx")

    def Giris_bilgileri_alma(self):
        self.Sp_Id=""
        self.Ozel_alan1=""
        self.Ozel_alan2=""
        self.Ozel_alan3=""
        self.selenium.open_available_browser("https://dev.smartportal.com.tr/")
        self.selenium.input_text(locator="xpath://*[@autocomplete='email']",text="system@smartportal.com")
        self.selenium.input_text(locator="xpath://*[@autocomplete='current-password']",text="smartportal")
        self.selenium.click_button(locator="xpath://*[@class='mat-focus-indicator w-100 mat-raised-button mat-button-base mat-primary']")
        self.selenium.wait_until_element_is_visible(locator="xpath://*[@routerlink='/isyeri']")
        self.selenium.click_element(locator="xpath://*[@routerlink='/isyeri']")
        self.selenium.wait_until_element_is_visible(locator="xpath://*[@id='main-sidenav']/div/app-isyeri-list-page/div/app-isyeri-list-table/div/table/tbody")
        row_num2=self.selenium.get_element_count(locator="xpath://*[@id='main-sidenav']/div/app-isyeri-list-page/div/app-isyeri-list-table/div/table/tbody/tr")
        
        
        for j in range(1,row_num2+1):
            
                self.sirket2=self.selenium.get_webelement(locator=("xpath://*[@id='main-sidenav']/div/app-isyeri-list-page/div/app-isyeri-list-table/div/table/tbody/tr[{}]/td[3]".format(j)))
                if self.sirket ==self.sirket2.get_attribute("innerHTML").strip():
                    self.Sp_Id=self.selenium.get_webelement(locator=("xpath://*[@id='main-sidenav']/div/app-isyeri-list-page/div/app-isyeri-list-table/div/table/tbody/tr[{}]/td[2]".format(j))).get_attribute("innerHTML").strip()
                    print("ıd:",self.Sp_Id)
                    
                    self.Ozel_alan1=self.selenium.get_webelement(locator=("xpath://*[@id='main-sidenav']/div/app-isyeri-list-page/div/app-isyeri-list-table/div/table/tbody/tr[{}]/td[4]".format(j))).get_attribute("innerHTML").strip()
                    print("1",self.Ozel_alan1)
                    
                    
                    self.Ozel_alan2=self.selenium.get_webelement(locator=("xpath://*[@id='main-sidenav']/div/app-isyeri-list-page/div/app-isyeri-list-table/div/table/tbody/tr[{}]/td[5]".format(j))).get_attribute("innerHTML").strip()
                    print("2",self.Ozel_alan2)
                   
                    self.Ozel_alan3=self.selenium.get_webelement(locator=("xpath://*[@id='main-sidenav']/div/app-isyeri-list-page/div/app-isyeri-list-table/div/table/tbody/tr[{}]/td[6]".format(j))).get_attribute("innerHTML").strip()
                    print("3",self.Ozel_alan3)
                    
                    self.selenium.click_element(locator="xpath://*[@id='main-sidenav']/div/app-isyeri-list-page/div/app-isyeri-list-table/div/table/tbody/tr[{}]/td[7]/button[1]".format(j))
                    break
                    
                else:
                    continue
            
            
        self.selenium.wait_until_element_is_visible(locator="xpath://*[@class='mat-cell cdk-cell cdk-column-kullaniciAdi mat-column-kullaniciAdi ng-star-inserted']")

        self.kullanici_ad=self.selenium.get_webelement(locator="xpath://*[@class='mat-cell cdk-cell cdk-column-kullaniciAdi mat-column-kullaniciAdi ng-star-inserted']")
        self.kullanici_ad=self.kullanici_ad.get_attribute('innerHTML').strip()
        
        self.isyeri_kodu=self.selenium.get_webelement(locator="xpath://*[@class='mat-cell cdk-cell cdk-column-isyeriKodu mat-column-isyeriKodu ng-star-inserted']")
        self.isyeri_kodu=self.isyeri_kodu.get_attribute('innerHTML').strip()
        
        
        self.sistem_sifre=self.selenium.get_webelement(locator="xpath://*[@class='mat-cell cdk-cell cdk-column-sistemSifresi mat-column-sistemSifresi ng-star-inserted']")
        self.sistem_sifre=self.sistem_sifre.get_attribute('innerHTML').strip()
        
        self.isyeri_sifre=self.selenium.get_webelement(locator="xpath://*[@class='mat-cell cdk-cell cdk-column-isyeriSifresi mat-column-isyeriSifresi ng-star-inserted']")
        self.isyeri_sifre=self.isyeri_sifre.get_attribute('innerHTML').strip()
        
        self.ünvan=self.selenium.get_webelement(locator="xpath://*[@id='main-sidenav']/div/app-isyeri-view/div/div/app-view-item-text[3]/div/div[2]")
        self.ünvan=self.ünvan.get_attribute('innerHTML').strip()
        
        self.sicil_no=self.selenium.get_webelement(locator="xpath://*[@id='main-sidenav']/div/app-isyeri-view/div/div/app-view-item-text[1]/div/div[2]")
        self.sicil_no=self.sicil_no.get_attribute('innerHTML')
        self.ayrac=self.sicil_no[23:26] 
        self.selenium.close_browser()
        
        
    def Giris_bilgileri_yazma(self):    
        self.selenium.open_available_browser("https://uyg.sgk.gov.tr/IsverenSistemi")
        for attempts in range(0,3):
            self.selenium.input_text(locator="xpath://*[@id='kullaniciIlkKontrollerGiris_username']",text=self.kullanici_ad)
            self.selenium.input_text(locator="xpath://*[@name='isyeri_kod']",text=self.isyeri_kodu)
            self.selenium.input_text(locator="xpath://*[@ name='password']",text=self.sistem_sifre)
            self.selenium.input_text(locator="xpath://*[@name='isyeri_sifre']",text=self.isyeri_sifre)
        
            #img=self.selenium.get_webelement(locator="xpath://*[@id='guvenlik_kod']")
            #url=img.get_attribute("src")
            #ret_code, ret_log, chaptca_text = self.chaptcasolve(url, 5, 5)
            #self.logger.info(f"Chaptca; : ret_code: {ret_code}, ret_log: {ret_log}, chaptca_text: {chaptca_text}")

            #if ret_code == 0:
            #    self.selenium.execute_javascript(f"document.getElementById('kullaniciIlkKontrollerGiris_isyeri_guvenlik').setAttribute('value', '{chaptca_text}')")
            #    self.selenium.click_button("Giriş Yap")
            #    break
            #else: 
             #   if attempts ==2:
              #      self.selenium.execute_javascript(f"document.getElementById('kullaniciIlkKontrollerGiris_isyeri_guvenlik').setAttribute('value', 'Hata! Gir!')")
              #  else:
               #     continue
   
   
    
    def sicil(self):

        self.selenium.switch_window('İşveren Sistemi',1000)
       
        self.selenium.maximize_browser_window()
        self.selenium.wait_until_element_is_visible(locator="xpath://*[@id='menuForm:panelMenuId']/h3[2]/span")
        self.selenium.click_element(locator="xpath://*[@id='menuForm:panelMenuId']/h3[2]/span")
        self.selenium.wait_until_element_is_visible(locator="xpath://*[@id='menuForm:anaMenu1subItem0']")
        self.selenium.click_element(locator="xpath://*[@id='menuForm:anaMenu1subItem0']")
        self.selenium.select_frame(locator="xpath://*[@id='pencereLinkIdYeni']")

        self.selenium.wait_until_element_is_visible(locator="xpath://*[@id='secim']")
        
        self.selenium.click_element(locator="xpath://*[@id='secim']")

        self.selenium.wait_until_element_is_visible(locator="xpath://*[@id='secim']/option[4]")
        
        self.selenium.click_element(locator="xpath://*[@id='secim']/option[4]")  
              
               
        self.selenium.wait_until_element_is_visible(locator="xpath://*[@id='sec']")
        
        
        row_num3=self.selenium.get_element_count(locator="xpath://*[@id='sec']/option")
        st=row_num3
        end=row_num3
        for opt in range(1,row_num3+1):
            datee=self.selenium.get_webelement(locator="xpath://*[@id='sec']/option[{}]".format(opt))
            if self.start_date == str(datee.get_attribute("innerHTML")).strip() :
                st=opt
            if self.end_date == str(datee.get_attribute("innerHTML")).strip():
                end=opt
            
            else:
                continue

        print("end",end)
        print("st",st)
        
        self.excel.create_workbook("asgari.xlsx")

        self.excel.set_cell_value(1, "A", "Şirket")
        self.excel.set_cell_value(1, "B", "Yıl" )
        self.excel.set_cell_value(1, "F", "Faydalınalınan gün sayısı")
        self.excel.set_cell_value(1, "D", "Destek Tutarı")
        self.excel.set_cell_value(1, "E", "Tahsilat tarihi" )
        self.excel.set_cell_value(1, "C", "Ay")
        self.excel.set_cell_value(1, "G", "Sp_Id")
        self.excel.set_cell_value(1, "H", "Ozel_Alan_1")
        self.excel.set_cell_value(1, "I", "Ozel_Alan_2")
        self.excel.set_cell_value(1, "J", "Ozel_Alan_3")
        self.excel.save_workbook("asgari.xlsx")
        self.excel.close_workbook()        
        for x in range(st,end+1):
    
            self.tarih=self.selenium.get_webelement(locator="xpath://*[@id='sec']/option[{}]".format(x)).get_attribute("innerHTML").strip()
            self.selenium.click_element(locator="xpath://*[@id='sec']/option[{}]".format(x))

            self.selenium.click_button("Listele")
            row_num=(self.selenium.get_element_count(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr"))
            print("l",row_num)
            if row_num==0:
                self.islem_tarihi=0
                self.faydalınalınan_gün_sayisi=0
                self.destek_tutari=0
                self.tahsilat_tarihi=0
                    
                
            else:
                for a in range(2,row_num+1):
                    self.yil=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[{}]/td[2]/p".format(a)).get_attribute("innerHTML").strip()
                    self.ay=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[{}]/td[3]/p".format(a)).get_attribute("innerHTML").strip()
                    self.faydalınalınan_gün_sayisi=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[{}]/td[4]/p".format(a)).get_attribute("innerHTML").strip()
                    self.destek_tutari=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[{}]/td[5]/p".format(a)).get_attribute("innerHTML").strip()
                    self.tahsilat_tarihi=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[{}]/td[6]/p".format(a)).get_attribute("innerHTML").strip()
                
                
                    self.excel.open_workbook("asgari.xlsx")
                    self.excel.set_cell_value(a, "A", self.sirket)
                    self.excel.set_cell_value(a, "B", self.yil)
                    self.excel.set_cell_value(a, "F", self.faydalınalınan_gün_sayisi)
                    self.excel.set_cell_value(a, "D",self.destek_tutari)
                    self.excel.set_cell_value(a, "E", self.tahsilat_tarihi )
                    self.excel.set_cell_value(a, "C", self.ay)
                    self.excel.set_cell_value(a, "G", self.Sp_Id)
                    self.excel.set_cell_value(a, "H", self.Ozel_alan1)
                    self.excel.set_cell_value(a, "I", self.Ozel_alan2)
                    self.excel.set_cell_value(a, "J", self.Ozel_alan3)
                    self.excel.save_workbook("asgari.xlsx")
                    self.excel.close_workbook()
        
    def run(self):
        Task.Giris_bilgileri_alma(self)
        Task.Giris_bilgileri_yazma(self)
        Task.sicil(self)
       

  

        
    


        
                
if __name__ == "__main__":
    #Task.tarih(self)
    Task().run()
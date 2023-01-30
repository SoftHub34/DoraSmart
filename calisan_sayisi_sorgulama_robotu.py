#Gerekli kütüphaneler importlandı. çalışsan sayısı sorgulama robotu ilk demo yapılan ve demo için hazırlanan robot.
from task_base import TaskBase #her robot için gerekli olan lisansları kontrol eden çalışmasını kontrol eden task_base kodu importlandı
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
import time
import sys
import json

#class yapısı oluşturuldu.Gerekli parametreler oluşturuldu. İçerisine try-except bloğu ile çalışıp çalşmama durumu kontrol altına alındı
class Task(TaskBase):
    def __init__(self, task_parameters):
        try:
            super().__init__(task_parameters)
            self.logger = self.get_logger()
            self.logger.info(f"Started.")
            task_parameters = json.loads(task_parameters)

            self.portal_url = task_parameters['portal_url']
            self.portal_user_mail = task_parameters['portal_user_mail']
            self.portal_user_password = task_parameters['portal_user_password']
            self.selected_company = task_parameters['selected_company']
            self.sgk_url = task_parameters['sgk_url']
            self.donem_baslangic = task_parameters['donem_baslangic']
            self.donem_bitis = task_parameters['donem_bitis']

            self.selenium = Selenium()
            self.excel = Files()
            self.step_no = 1
            self.logger.info(f"Finished.")
        except Exception as e:
            self.logger.info(f"Failed. Error: {e}")        


    def execute(self):
        return_code = 0
        try:
            self.logger.info(f"Started.")
            status = self.get_status(True)
            status_log = ''
            self.execute_begin()

            self.Giris_bilgileri_alma()
            self.Giris_bilgileri_yazma()
            self.sicil()

        except Exception as e:
            self.logger.info(f"Failed. Error: {e}")        
            status = self.get_status(False)
            status_log = f"{e}"
            return_code = 1
            
        self.execute_end(status, status_log)
        self.logger.info(f"Finished. return_code: {return_code}")        
        return return_code

#Robot kodunun çalışmasının başlangıcında sgk sitesine girebilmek için verilerin çekilmesi gerekmektedir verilerin alınacağı siteye yönlendiren verileri toplayan fonksiyon oluşturuldu.
    def Giris_bilgileri_alma(self):
        self.logger.info(f"Started.")
        self.Sp_Id="" #Vereceğimiz excel raporunda istenen ve giriş bilgilerini aldığımız smartportal adresinde olan verileri de topluyoruz.
        self.Ozel_alan1=""#Vereceğimiz excel raporunda istenen ve giriş bilgilerini aldığımız smartportal adresinde olan verileri de topluyoruz.
        self.Ozel_alan2=""#Vereceğimiz excel raporunda istenen ve giriş bilgilerini aldığımız smartportal adresinde olan verileri de topluyoruz.
        self.Ozel_alan3=""#Vereceğimiz excel raporunda istenen ve giriş bilgilerini aldığımız smartportal adresinde olan verileri de topluyoruz.
        self.selenium.open_available_browser(self.portal_url)
        self.selenium.input_text(locator="xpath://*[@autocomplete='email']",text=self.portal_user_mail)
        self.selenium.input_text(locator="xpath://*[@autocomplete='current-password']",text=self.portal_user_password)
        self.selenium.click_button(locator="xpath://*[@class='mat-focus-indicator w-100 mat-raised-button mat-button-base mat-primary']")
        self.selenium.wait_until_element_is_visible(locator="xpath://*[@routerlink='/isyeri']")
        self.selenium.click_element(locator="xpath://*[@routerlink='/isyeri']")
        self.selenium.wait_until_element_is_visible(locator="xpath://*[@id='main-sidenav']/div/app-isyeri-list-page/div/app-isyeri-list-table/div/table/tbody")

        row_num2=self.selenium.get_element_count(locator="xpath://*[@id='main-sidenav']/div/app-isyeri-list-page/div/app-isyeri-list-table/div/table/tbody/tr")
        
        
        for j in range(1,row_num2+1):
            
                self.sirket2=self.selenium.get_webelement(locator=("xpath://*[@id='main-sidenav']/div/app-isyeri-list-page/div/app-isyeri-list-table/div/table/tbody/tr[{}]/td[3]".format(j)))
                if self.selected_company.strip() ==self.sirket2.get_attribute("innerHTML").strip():
                    self.Sp_Id=self.selenium.get_webelement(locator=("xpath://*[@id='main-sidenav']/div/app-isyeri-list-page/div/app-isyeri-list-table/div/table/tbody/tr[{}]/td[2]".format(j))).get_attribute("innerHTML").strip()
                    self.logger.info(f"Sp_Id: {self.Sp_Id}")
                    
                    self.Ozel_alan1=self.selenium.get_webelement(locator=("xpath://*[@id='main-sidenav']/div/app-isyeri-list-page/div/app-isyeri-list-table/div/table/tbody/tr[{}]/td[4]".format(j))).get_attribute("innerHTML").strip()
                    self.logger.info(f"Ozel_alan1: {self.Ozel_alan1}")
                    
                    self.Ozel_alan2=self.selenium.get_webelement(locator=("xpath://*[@id='main-sidenav']/div/app-isyeri-list-page/div/app-isyeri-list-table/div/table/tbody/tr[{}]/td[5]".format(j))).get_attribute("innerHTML").strip()
                    self.logger.info(f"Ozel_alan2: {self.Ozel_alan2}")
                
                    self.Ozel_alan3=self.selenium.get_webelement(locator=("xpath://*[@id='main-sidenav']/div/app-isyeri-list-page/div/app-isyeri-list-table/div/table/tbody/tr[{}]/td[6]".format(j))).get_attribute("innerHTML").strip()
                    self.logger.info(f"Ozel_alan3: {self.Ozel_alan3}")
                    
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
        
        self.selenium.wait_until_element_is_visible(locator="xpath://*[@id='main-sidenav']/div/app-isyeri-view/div/div/app-view-item-text[3]/div/div[2]", timeout=20)
        self.unvan=self.selenium.get_webelement(locator="xpath://*[@id='main-sidenav']/div/app-isyeri-view/div/div/app-view-item-text[3]/div/div[2]")
        self.unvan=self.unvan.get_attribute('innerHTML').strip()
        
        self.sicil_no=self.selenium.get_webelement(locator="xpath://*[@id='main-sidenav']/div/app-isyeri-view/div/div/app-view-item-text[1]/div/div[2]")
        self.sicil_no=self.sicil_no.get_attribute('innerHTML')
        self.ayrac=self.sicil_no[23:26] 
        self.selenium.close_browser()
        self.logger.info(f"Finished.")
        
    #alınan giriş bilgileri ile sgk sitesine giriş yapılır. bilgiler ekrana yazılarak işleme devam edilir.(Captcha kodu denenir fakat çözülemezse kullanıcıya manuel girmesi istenebilir.)    
    def Giris_bilgileri_yazma(self):    
        self.logger.info(f"Started.")
        self.selenium.open_available_browser(self.sgk_url)
        self.selenium.input_text(locator="xpath://*[@id='kullaniciIlkKontrollerGiris_username']",text=self.kullanici_ad)
        self.selenium.input_text(locator="xpath://*[@name='isyeri_kod']",text=self.isyeri_kodu)
        self.selenium.input_text(locator="xpath://*[@ name='password']",text=self.sistem_sifre)
        self.selenium.input_text(locator="xpath://*[@name='isyeri_sifre']",text=self.isyeri_sifre)
        
        img=self.selenium.get_webelement(locator="xpath://*[@id='guvenlik_kod']")
        url=img.get_attribute("src")
        ret_code, ret_log, chaptca_text = self.solve_chaptca(url, 5, 5)
        self.logger.info(f"Chaptca; : ret_code: {ret_code}, ret_log: {ret_log}, chaptca_text: {chaptca_text}")

        if ret_code == 0:
            self.selenium.execute_javascript(f"document.getElementById('kullaniciIlkKontrollerGiris_isyeri_guvenlik').setAttribute('value', '{chaptca_text}')")
            self.selenium.click_button("Giriş Yap")
        else:
            self.selenium.execute_javascript(f"document.getElementById('kullaniciIlkKontrollerGiris_isyeri_guvenlik').setAttribute('value', 'Hata! Gir!')")
            #self.selenium.driver.execute_script("alert('Chaptca çözülemedi! Lütfen giriş yapın...')")
        if ret_code != 0:
            self.selenium.input_text(locator="xpath://*[@name='isyeri_kod']",text=self.isyeri_kodu)
            self.selenium.input_text(locator="xpath://*[@ name='password']",text=self.sistem_sifre)
            
        self.selenium.wait_until_page_does_not_contain_element(locator="xpath://*[@id='guvenlik_kod']", timeout=180)
        self.logger.info(f"Finished.")
   
    #siteye girilmesinin ardından çıktı olarak bizden istenen verilerin hepsinin toplanması için sicil fonksiyonu oluştrulur.
    #verilerin toplanmasındaki şartlar aşağıda belirtilmiştir.
    #Açılan ekranda işçi sayısı tek satır ise aracı sıra noya göre hareket edilir ve işleme devam edilir.
    #Çalışan sayısı çekilmek istenen işyerinin sicil numarası işaretli yerde yani işyeri kartı aracı sıra no '000' ise,
    #İlk sırada yer alan başında herhangi bir rakam yazmayan satırın karşısındaki sayı işyeri çalışan sayısına yazılır.
    #Toplamda bu sayı çıkartılır ve alt işveren sayısına yazılır. Diğer mükerrer bildrim ve tekilleştirilmiş sayı direkt alınır.
    #Eğer 036 olarak işaretlenen gibi aracı sıranosu var ise bu satırın karşısındaki  sayı işyeri çalışan sayısına yazılır. 
    #Başında sayı olmayan satır sayısı üst dosya çalışan sayısına yazılır. Toplam sayıdan üst dosya çalışan sayısı çıkartılır ve alt işveren çalışan sayısı bulunur
    #diğer mükerrer bildrim ve tekilleştirilmiş sayı direkt alınır.
    def sicil(self):
        self.logger.info(f"Started.")
        self.is_yeri_calisan_sayisi=0#fonksiyonla sıfıra eşitlenir ve sonrasında toplanan veriler eklenir.
        self.ust_dosya_calisan_sayisi=0#fonksiyonla sıfıra eşitlenir ve sonrasında toplanan veriler eklenir.
        self.alt_isveren_calisan_sayisi=0#fonksiyonla sıfıra eşitlenir ve sonrasında toplanan veriler eklenir.
        self.mukerrer_bildirim=0#fonksiyonla sıfıra eşitlenir ve sonrasında toplanan veriler eklenir.
        self.tekillestirilmis_toplam_sayi=0#fonksiyonla sıfıra eşitlenir ve sonrasında toplanan veriler eklenir.
        self.selenium.switch_window('İşveren Sistemi',1000)
    
        self.selenium.maximize_browser_window()
        self.selenium.wait_until_element_is_visible(locator="xpath://*[@id='menuForm:panelMenuId']/h3[2]/span")

        self.selenium.click_element(locator="xpath://*[@id='menuForm:panelMenuId']/h3[2]/span")
        self.selenium.wait_until_element_is_visible(locator="xpath://*[@id='menuForm:anaMenu1subMenu2']/span[1]")
        self.selenium.click_element(locator="xpath://*[@id='menuForm:anaMenu1subMenu2']/span[1]")
        self.selenium.wait_until_element_is_visible(locator="xpath://*[@id='menuForm:anaMenu1subMenu2subItem1']/span")
        
        self.selenium.click_element(locator="xpath://*[@id='menuForm:anaMenu1subMenu2subItem1']/span")
        time.sleep(10)
        
        self.selenium.select_frame(locator="xpath://*[@id='pencereLinkIdYeni']")
            
        self.selenium.click_element(locator="xpath://*[@id='ortalamaCalisan_sent_donem_yil_ay_index']")
    
    
        row_num3=self.selenium.get_element_count(locator="xpath://*[@id='ortalamaCalisan_sent_donem_yil_ay_index']/option")
        self.logger.info(f"Donem count: {row_num3}")

        for opt in range(1,row_num3+1):
            datee=self.selenium.get_webelement(locator="xpath://*[@id='ortalamaCalisan_sent_donem_yil_ay_index']/option[{}]".format(opt))
            if self.donem_baslangic == str(datee.get_attribute("innerHTML")).strip() :
                st=opt
            if self.donem_bitis == str(datee.get_attribute("innerHTML")).strip():
                end=opt
            else:
                continue

    #excel dosyası ve sütunlar oluşturulur.
        self.excel.create_workbook("calışan_sayısı.xlsx")
        self.excel.set_cell_value(1, "A", "Şirket")
        self.excel.set_cell_value(1, "B", "is_yeri_calisan_sayisi" )
        self.excel.set_cell_value(1, "C", "üst_dosya_calisan_sayisi")
        self.excel.set_cell_value(1, "D", "alt_isveren_calisan_sayisi")
        self.excel.set_cell_value(1, "E", "mükerrer_bildirim" )
        self.excel.set_cell_value(1, "F", "tekillestirilmis_toplam_sayi")
        self.excel.set_cell_value(1, "G", "Sp_Id")
        self.excel.set_cell_value(1, "H", "Ozel_Alan_1")
        self.excel.set_cell_value(1, "I", "Ozel_Alan_2")
        self.excel.set_cell_value(1, "J", "Ozel_Alan_3")
        self.excel.set_cell_value(1, "K", "Tarih")
        for x in range(st,end+1):
    
            self.tarih=self.selenium.get_webelement(locator="xpath://*[@id='ortalamaCalisan_sent_donem_yil_ay_index']/option[{}]".format(x)).get_attribute("innerHTML").strip()
            self.selenium.click_element(locator="xpath://*[@id='ortalamaCalisan_sent_donem_yil_ay_index']/option[{}]".format(x))

            self.selenium.click_button("Listele")

            row_num=(self.selenium.get_element_count(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr"))

            if self.ayrac== "000":
                self.is_yeri_calisan_sayisi=0
                self.ust_dosya_calisan_sayisi=0
                self.alt_isveren_calisan_sayisi=0
                self.mukerrer_bildirim=0
                self.tekillestirilmis_toplam_sayi=0
                if row_num !=0:
                    self.unvan2=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[2]/td[1]/p")
                    if self.unvan ==self.unvan2.get_attribute("innerHTML") :
                        self.unvandeger=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[2]/td[3]/p")
                        self.is_yeri_calisan_sayisi += int(self.unvandeger.get_attribute("innerHTML"))
                
                    self.ust_dosya_calisan_sayisi += 0

            
                    self.AltIsverenCalisanSayisi2 = self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[3]/th/p")
                    if "TOPLAM"  == self.AltIsverenCalisanSayisi2.get_attribute("innerHTML"): 
                        self.alt_isveren_calisan_sayisi_deger=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[3]/td/p")
                        self.alt_isveren_calisan_sayisi+= int(self.alt_isveren_calisan_sayisi_deger.get_attribute("innerHTML"))

                    self.mukerrer_bildirim+=0

            
                    self.tekillestirilmis_toplam_sayi+=self.alt_isveren_calisan_sayisi

                else:
                    pass
        
                step_value = {}
                self.excel.set_cell_value(x, "B", self.is_yeri_calisan_sayisi )
                step_value['isyeri_calisan_sayisi'] = self.is_yeri_calisan_sayisi
                self.excel.set_cell_value(x, "C", self.ust_dosya_calisan_sayisi)
                step_value['ust_dosya_calisan_sayisi'] = self.ust_dosya_calisan_sayisi
                self.excel.set_cell_value(x, "D", self.alt_isveren_calisan_sayisi)
                step_value['alt_isveren_calisan_sayisi'] = self.alt_isveren_calisan_sayisi
                self.excel.set_cell_value(x, "E", self.mukerrer_bildirim )
                step_value['mukerrer_bildirim'] = self.mukerrer_bildirim
                self.excel.set_cell_value(x, "F", self.tekillestirilmis_toplam_sayi)
                step_value['tekillestirilmis_toplam_sayi'] = self.tekillestirilmis_toplam_sayi
                self.excel.set_cell_value(x, "A", self.unvan)
                step_value['unvan'] = self.unvan
                self.excel.set_cell_value(x, "G", self.Sp_Id)
                step_value['sp_id'] = self.Sp_Id
                self.excel.set_cell_value(x, "H", self.Ozel_alan1)
                step_value['ozel_alan1'] = self.Ozel_alan1
                self.excel.set_cell_value(x, "I", self.Ozel_alan2)
                step_value['ozel_alan1'] = self.Ozel_alan2
                self.excel.set_cell_value(x, "J", self.Ozel_alan3)
                step_value['ozel_alan3'] = self.Ozel_alan3
                self.excel.set_cell_value(x, "K", self.tarih)
                step_value['tarih'] = self.tarih
                self.add_trx_step(self.step_no, step_value)
                self.step_no += 1
                    
            if self.ayrac != "000":
                self.is_yeri_calisan_sayisi=0
                self.ust_dosya_calisan_sayisi=0
                self.alt_isveren_calisan_sayisi=0
                self.mukerrer_bildirim=0
                self.tekillestirilmis_toplam_sayi=0

                if row_num !=0:
                    self.mukerrer_bildirim2=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[{}]/td[1]/p".format(row_num-1))            
                    if "Asıl İşveren ve Alt İşverence Aynı Kişi İçin Yapılan Mükerrer Bildirim" ==self.mukerrer_bildirim2.get_attribute("innerHTML").strip():
                        self.mukerrer_bildirim_deger=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[{}]/td[3]/p".format(row_num-1))
                        self.mukerrer_bildirim += int(self.mukerrer_bildirim_deger.get_attribute("innerHTML"))
                    else:
                        break 
                
                    self.ust_dosya_calisan_sayisi2=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[2]/td[3]/p")
                    self.ust_dosya_calisan_sayisi += int(self.ust_dosya_calisan_sayisi2.get_attribute("innerHTML"))
                
                    self.AltIsverenCalisanSayisi2 = self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[{}]/th/p".format(row_num-2))
                    if "TOPLAM"  == self.AltIsverenCalisanSayisi2.get_attribute("innerHTML"): 
                        self.alt_isveren_calisan_sayisi_deger=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[{}]/td/p".format(row_num-2))
                        self.alt_isveren_calisan_sayisi+= int(self.alt_isveren_calisan_sayisi_deger.get_attribute("innerHTML"))
    
                    self.tekillestirilmis_toplam_sayi2=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[{}]/th/p".format(row_num))
                    if "ASIL VE ALT İŞVERENLERCE BİLDİRİLEN AYNI SİGORTALININ TEK SAYILDIĞI TOPLAM" ==self.tekillestirilmis_toplam_sayi2.get_attribute("innerHTML"):
                        self.tekillestirilmis_toplam_sayi += (self.alt_isveren_calisan_sayisi-self.mukerrer_bildirim)
                    
                    for i in range(2,row_num-3):
                        self.unvan2=self.selenium.get_webelement(locator=("xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[{}]/td[1]/p".format(i)))
                        if self.ayrac ==self.unvan2.get_attribute("innerHTML")[0:3]:
                            self.unvandeger=self.selenium.get_webelement(locator="xpath://*[@id='contentContainer']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table[2]/tbody/tr[{}]/td[3]/p".format(i))
                            self.is_yeri_calisan_sayisi += int(self.unvandeger.get_attribute("innerHTML"))
                        else:
                            continue
                            
                else:
                    pass            
                            
                step_value = {}
                self.excel.set_cell_value(x, "B", self.is_yeri_calisan_sayisi )
                step_value['isyeri_calisan_sayisi'] = self.is_yeri_calisan_sayisi
                self.excel.set_cell_value(x, "C", self.ust_dosya_calisan_sayisi)
                step_value['ust_dosya_calisan_sayisi'] = self.ust_dosya_calisan_sayisi
                self.excel.set_cell_value(x, "D", self.alt_isveren_calisan_sayisi)
                step_value['alt_isveren_calisan_sayisi'] = self.alt_isveren_calisan_sayisi
                self.excel.set_cell_value(x, "E", self.mukerrer_bildirim )
                step_value['mukerrer_bildirim'] = self.mukerrer_bildirim
                self.excel.set_cell_value(x, "F", self.tekillestirilmis_toplam_sayi)
                step_value['tekillestirilmis_toplam_sayi'] = self.tekillestirilmis_toplam_sayi
                self.excel.set_cell_value(x, "A", self.unvan)
                step_value['unvan'] = self.unvan
                self.excel.set_cell_value(x, "G", self.Sp_Id)
                step_value['sp_id'] = self.Sp_Id
                self.excel.set_cell_value(x, "H", self.Ozel_alan1)
                step_value['ozel_alan1'] = self.Ozel_alan1
                self.excel.set_cell_value(x, "I", self.Ozel_alan2)
                step_value['ozel_alan2'] = self.Ozel_alan2
                self.excel.set_cell_value(x, "J", self.Ozel_alan3)
                step_value['ozel_alan3'] = self.Ozel_alan3
                self.excel.set_cell_value(x, "K", self.tarih)
                step_value['tarih'] = self.tarih
                self.add_trx_step(self.step_no, step_value)
                self.step_no += 1
        #excel dosyası eklenen verilerle kaydedilir. ve kapatılır.        
        self.excel.save_workbook("calışan_sayısı.xlsx")
        self.excel.close_workbook()


def main():
    return Task(sys.argv[1].replace("'", '"')).execute()
            
if __name__ == "__main__":
    main()

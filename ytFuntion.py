from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from openpyxl import workbook ,load_workbook
import os ,time ,random ,ytFuntion

testday_File = time.strftime("%y_%m_%d")
testday_Time  = time.strftime("%y_%m_%d_%H_%M_%S")

funtion_error = []
funtion_count_Png = 1

class account_Setting():
    def __init__(self ,username = "" ,password = "" ,safe_Password = ""):
        self.username = str(username).strip()
        self.password = str(password).strip()
        self.safe_Password = str(safe_Password).strip()

class test_web():
    def __init__(self ,web_Driver):
        self.web_Driver = web_Driver

    def period_Confirm(self):
        try:
            #self.element_Click("//span[.='确定']" ,8)
            self.web_Driver.find_element_by_xpath("//span[.='确定']").click()
        except:
            pass

    def web_item(self):
        return self.web_Driver.find_element_by_css_selector("ul[class='betFilter']").find_elements_by_tag_name('li') #取得item,固定寫法

    def web_item_click(self ,i):
        self.period_Confirm()
        if self.web_item()[i].text != "二同号单选":
            self.web_item()[i].click()

    def web_Page(self):
        return self.web_Driver.find_element_by_css_selector("ul[class='betNav fix']").find_elements_by_tag_name('li') #取得分頁,固定寫法

    def web_Page_click(self ,i ,element_Text = "" ,link_type = None):
        self.period_Confirm()
        self.web_Page()[i].click()
        if i >= 5 and i < len(self.web_Page()) and len(self.web_Page()) > 6:
            self.element_Click(element_Text ,link_type)

    def save_Png(self ,save_Text = None ,drop_Down_count = "" ,donot_Save = ""):
        if save_Text == None or str(donot_Save) != "":
            return
        global funtion_error ,funtion_count_Png  #全域變數被當成區域變數的解法
        web_Height = self.web_Driver.execute_script("return document.body.scrollHeight")
        web_Position_y ,i ,drop_Down = 0 ,1 ,1
        while i <= drop_Down:
            try:
                self.web_Driver.execute_script("window.scroll(0, "+ str(web_Position_y) +");")
                sleep(1)
                self.period_Confirm()
                self.web_Driver.save_screenshot(testday_File + "/" + str(testday_Time) + "_" + str(funtion_count_Png) + "_" + str(save_Text) + ".png")
                funtion_count_Png += 1
                web_Position_y += 600
                i += 1
                if drop_Down_count != "":
                    drop_Down = int(drop_Down_count)
                else:
                    drop_Down = (self.web_Driver.execute_script("return document.body.scrollHeight") / 600) + 1
            except:
                funtion_error.append(save_Text + str(funtion_count_Png) + "_NG")
                return funtion_error

    def element_Click(self ,element_Text = "" ,link_type = None ,delay_Time = 0):
        global funtion_error #全域變數被當成區域變數的解法
        try:
            link_type = int(str(link_type).strip())
            element_Text = str(element_Text).strip()
            delay_Time = int(str(delay_Time).strip())
            if link_type == 1:
                WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_element_located((By.ID,element_Text)))
                if delay_Time != 0:
                    sleep(delay_Time)
                self.period_Confirm()
                self.web_Driver.find_element_by_id(element_Text).click()
            elif link_type == 2:
                WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME,element_Text)))
                if delay_Time != 0:
                    sleep(delay_Time)
                self.period_Confirm()
                self.web_Driver.find_element_by_class_name(element_Text).click()
            elif link_type == 3:
                WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_element_located((By.LINK_TEXT,element_Text)))
                if delay_Time != 0:
                    sleep(delay_Time)
                self.period_Confirm()
                self.web_Driver.find_element_by_link_text(element_Text).click()
            elif link_type == 4:
                WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_element_located((By.PARTIAL_LINK_TEXT,element_Text)))
                if delay_Time != 0:
                    sleep(delay_Time)
                self.period_Confirm()
                self.web_Driver.find_element_by_partial_link_text(element_Text).click()
            elif link_type == 5:
                WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_element_located((By.NAME,element_Text)))
                if delay_Time != 0:
                    sleep(delay_Time)
                self.period_Confirm()
                self.web_Driver.find_element_by_name(element_Text).click()
            elif link_type == 6:
                WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR,element_Text)))
                if delay_Time != 0:
                    sleep(delay_Time)
                self.period_Confirm()
                self.web_Driver.find_element_by_css_selector(element_Text).click()
            elif link_type == 7:
                WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_element_located((By.TAG_NAME,element_Text)))
                if delay_Time != 0:
                    sleep(delay_Time)
                self.period_Confirm()
                self.web_Driver.find_element_by_tag_name(element_Text).click()
            elif link_type == 8:
                WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_element_located((By.XPATH,element_Text)))
                if delay_Time != 0:
                    sleep(delay_Time)
                self.period_Confirm()
                self.web_Driver.find_element_by_xpath(element_Text).click()
            else:
                return funtion_error.append(element_Text + "_" + str(link_type) + "_ClickNG")
        except:
            funtion_error.append(element_Text + "_" + str(link_type) + "_NG")
            return funtion_error

    def elements_Click_one(self ,element_Text = "",link_type = None ,elements_num = 0 ,delay_Time = 0):
        global funtion_error #全域變數被當成區域變數的解法
        self.period_Confirm()
        try:
            link_type = int(str(link_type).strip())
            element_Text = str(element_Text).strip()
            elements_num = int(str(elements_num).strip()) -1
            if link_type == 1:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.ID,element_Text)))
                self.web_Driver.find_elements_by_id(element_Text)[elements_num].click()
            elif link_type == 2:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.CLASS_NAME,element_Text)))
                self.web_Driver.find_elements_by_class_name(element_Text)[elements_num].click()
            elif link_type == 3:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.LINK_TEXT,element_Text)))
                self.web_Driver.find_elements_by_link_text(element_Text)[elements_num].click()
            elif link_type == 4:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.PARTIAL_LINK_TEXT,element_Text)))
                self.web_Driver.find_elements_by_partial_link_text(element_Text)[elements_num].click()
            elif link_type == 5:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.NAME,element_Text)))
                self.web_Driver.find_elements_by_name(element_Text)[elements_num].click()
            elif link_type == 6:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR,element_Text)))
                self.web_Driver.find_elements_by_css_selector(element_Text)[elements_num].click()
            elif link_type == 7:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.TAG_NAME,element_Text)))
                self.web_Driver.find_elements_by_tag_name(element_Text)[elements_num].click()
            elif link_type == 8:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.XPATH,element_Text)))
                self.web_Driver.find_elements_by_xpath(element_Text)[elements_num].click()
            else:
                return funtion_error.append(element_Text + "_" + str(link_type) + "_ClickNG")
        except:
            funtion_error.append(element_Text + "_" + str(link_type) + "_NG")
            return funtion_error

    def elements_Click_all(self ,element_Text = "",link_type = None ,elements_num = 0 ,delay_Time = 0):
        global funtion_error #全域變數被當成區域變數的解法
        try:
            link_type = int(str(link_type).strip())
            element_Text = str(element_Text).strip()
            elements_num = int(str(elements_num).strip())
            if link_type == 1:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.ID,element_Text)))
                for i in range(elements_num):
                    self.period_Confirm()
                    self.web_Driver.find_elements_by_id(element_Text)[i].click()
            elif link_type == 2:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.CLASS_NAME,element_Text)))
                for i in range(elements_num):
                    self.period_Confirm()
                    self.web_Driver.find_elements_by_class_name(element_Text)[i].click()
            elif link_type == 3:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.LINK_TEXT,element_Text)))
                for i in range(elements_num):
                    self.period_Confirm()
                    self.web_Driver.find_elements_by_link_text(element_Text)[i].click()
            elif link_type == 4:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.PARTIAL_LINK_TEXT,element_Text)))
                for i in range(elements_num):
                    self.period_Confirm()
                    self.web_Driver.find_elements_by_partial_link_text(element_Text)[i].click()
            elif link_type == 5:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.NAME,element_Text)))
                for i in range(elements_num):
                    self.period_Confirm()
                    self.web_Driver.find_elements_by_name(element_Text)[i].click()
            elif link_type == 6:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR,element_Text)))
                for i in range(elements_num):
                    self.period_Confirm()
                    self.web_Driver.find_elements_by_css_selector(element_Text)[i].click()
            elif link_type == 7:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.TAG_NAME,element_Text)))
                for i in range(elements_num):
                    self.period_Confirm()
                    self.web_Driver.find_elements_by_tag_name(element_Text)[i].click()
            elif link_type == 8:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.XPATH,element_Text)))
                for i in range(elements_num):
                    self.period_Confirm()
                    self.web_Driver.find_elements_by_xpath(element_Text)[i].click()
            else:
                return funtion_error.append(element_Text + "_" + str(link_type) + "_ClickNG")
        except:
            funtion_error.append(element_Text + "_" + str(link_type) + "_NG")
            return funtion_error

    def elements(self ,element_Text = "",link_type = None):
        global funtion_error #全域變數被當成區域變數的解法
        self.period_Confirm()
        try:
            link_type = int(str(link_type).strip())
            element_Text = str(element_Text).strip()
            if link_type == 1:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.ID,element_Text)))
                return self.web_Driver.find_elements_by_id(element_Text)
            elif link_type == 2:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.CLASS_NAME,element_Text)))
                return self.web_Driver.find_elements_by_class_name(element_Text)
            elif link_type == 3:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.LINK_TEXT,element_Text)))
                return self.web_Driver.find_elements_by_link_text(element_Text)
            elif link_type == 4:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.PARTIAL_LINK_TEXT,element_Text)))
                return self.web_Driver.find_elements_by_partial_link_text(element_Text)
            elif link_type == 5:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.NAME,element_Text)))
                return self.web_Driver.find_elements_by_name(element_Text)
            elif link_type == 6:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR,element_Text)))
                return self.web_Driver.find_elements_by_css_selector(element_Text)
            elif link_type == 7:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.TAG_NAME,element_Text)))
                return self.web_Driver.find_elements_by_tag_name(element_Text)
            elif link_type == 8:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.XPATH,element_Text)))
                return self.web_Driver.find_elements_by_xpath(element_Text)
            else:
                funtion_error.append(element_Text + "_" + str(link_type) + "_get_NG")
                return funtion_error
        except:
            funtion_error.append(element_Text + "_" + str(link_type) + "_get_NG")
            return funtion_error

    def element(self ,element_Text = "",link_type = None):
        global funtion_error #全域變數被當成區域變數的解法
        self.period_Confirm()
        try:
            link_type = int(str(link_type).strip())
            element_Text = str(element_Text).strip()
            if link_type == 1:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.ID,element_Text)))
                return self.web_Driver.find_element_by_id(element_Text)
            elif link_type == 2:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.CLASS_NAME,element_Text)))
                return self.web_Driver.find_element_by_class_name(element_Text)
            elif link_type == 3:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.LINK_TEXT,element_Text)))
                return self.web_Driver.find_element_by_link_text(element_Text)
            elif link_type == 4:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.PARTIAL_LINK_TEXT,element_Text)))
                return self.web_Driver.find_element_by_partial_link_text(element_Text)
            elif link_type == 5:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.NAME,element_Text)))
                return self.web_Driver.find_element_by_name(element_Text)
            elif link_type == 6:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR,element_Text)))
                return self.web_Driver.find_element_by_css_selector(element_Text)
            elif link_type == 7:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.TAG_NAME,element_Text)))
                return self.web_Driver.find_element_by_tag_name(element_Text)
            elif link_type == 8:
                #WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.XPATH,element_Text)))
                return self.web_Driver.find_element_by_xpath(element_Text)
            else:
                funtion_error.append(element_Text + "_" + str(link_type) + "_ClickNG")
                return funtion_error
        except:
            funtion_error.append(element_Text + "_" + str(link_type) + "_NG")
            return funtion_error

    def element_Sendkeys(self ,element_Text = "" ,link_type = None ,delay_Time = 0 ,text = ""):
        global funtion_error #全域變數被當成區域變數的解法
        try:
            link_type = int(str(link_type).strip())
            element_Text = str(element_Text).strip()
            delay_Time = int(str(delay_Time).strip())
            text = str(text).strip()
            if link_type == 1:
                WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_element_located((By.ID,element_Text)))
                if delay_Time != 0:
                    sleep(delay_Time)
                self.period_Confirm()
                self.web_Driver.find_element_by_id(element_Text).send_keys(text)
            elif link_type == 2:
                WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME,element_Text)))
                if delay_Time != 0:
                    sleep(delay_Time)
                self.period_Confirm()
                self.web_Driver.find_element_by_class_name(element_Text).send_keys(text)
            elif link_type == 3:
                WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_element_located((By.LINK_TEXT,element_Text)))
                if delay_Time != 0:
                    sleep(delay_Time)
                self.period_Confirm()
                self.web_Driver.find_element_by_link_text(element_Text).send_keys(text)
            elif link_type == 4:
                WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_element_located((By.PARTIAL_LINK_TEXT,element_Text)))
                if delay_Time != 0:
                    sleep(delay_Time)
                self.period_Confirm()
                self.web_Driver.find_element_by_partial_link_text(element_Text).send_keys(text)
            elif link_type == 5:
                WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_element_located((By.NAME,element_Text)))
                if delay_Time != 0:
                    sleep(delay_Time)
                self.period_Confirm()
                self.web_Driver.find_element_by_name(element_Text).send_keys(text)
            elif link_type == 6:
                WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR,element_Text)))
                if delay_Time != 0:
                    sleep(delay_Time)
                self.period_Confirm()
                self.web_Driver.find_element_by_css_selector(element_Text).send_keys(text)
            elif link_type == 7:
                WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_element_located((By.TAG_NAME,element_Text)))
                if delay_Time != 0:
                    sleep(delay_Time)
                self.period_Confirm()
                self.web_Driver.find_element_by_tag_name(element_Text).send_keys(text)
            elif link_type == 8:
                WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_element_located((By.XPATH,element_Text)))
                if delay_Time != 0:
                    sleep(delay_Time)
                self.period_Confirm()
                self.web_Driver.find_element_by_xpath(element_Text).send_keys(text)
            else:
                return funtion_error.append(element_Text + "_" + str(link_type) + "_Send_key_NG")
        except:
            funtion_error.append(element_Text + "_" + str(link_type) + "_NG")
            return funtion_error

    def rebate(self ,element_Text = "" ,link_type = None ,element2_Text = "" ,link2_type = None ,rebate_Number = ""):
        link_type = int(str(link_type).strip())
        element_Text = str(element_Text).strip()
        link2_type = int(str(link2_type).strip())
        element2_Text = str(element2_Text).strip()
        rebate_Number = str(self.element(element_Text ,link_type).text)
        self.element_Sendkeys(element2_Text ,link2_type ,text = rebate_Number[0:-2])
        return rebate_Number[0:-2]
    
    def period_Detail(self):
        period_Detail = []
        while(True):
            WebDriverWait(self.web_Driver, 10).until(EC.visibility_of_all_elements_located((By.TAG_NAME,"td")))
            for i in range(len(self.elements("td" ,7))):
                try:
                    period_Detail.append(self.elements("td" ,7)[i].text)
                except:
                    pass
            try:
                self.web_Driver.find_element_by_xpath("//a[.='下一页']").click()
            except:
                break
        return period_Detail

    def speed_3_T_r(self ,element_Text = "" ,link_type = None ,max_Td = "0" ,max_Money = "0"):
        money = ["金額"]
        link_type = str(link_type).strip()
        element_Text = str(element_Text).strip()
        money_box = self.elements(element_Text ,link_type)
        max_Td = str(max_Td).strip()
        max_Money = str(max_Money).strip()
        if int(max_Td) == 0:
            money_box = money_box[0:-1]
        else:
            money_box = money_box[0:int(max_Td)]
        for i in range(1 ,len(money_box)):
            self.period_Confirm()
            money_box[i].clear()
            if int(max_Money) == 0:
                money.append(random.randint(0 ,99))
            else:
                money.append(int(max_Money))
            money_box[i].send_keys(str(money[i]))
        return money

    def speed_3_r(self ,element_Text = "" ,link_type = None ,max_Td = "0" ,max_Money = "0"):
        money = ["投注"]
        link_type = str(link_type).strip()
        element_Text = str(element_Text).strip()
        money_box = self.elements(element_Text ,link_type)
        max_Td = str(max_Td).strip()
        max_Money = str(max_Money).strip()
        if int(max_Td) == 0:
            money_box = money_box
        else:
            money_box = money_box[0:int(max_Td)]
        for i in range(len(money_box)):
            self.period_Confirm()
            money_box[i].clear()
            if int(max_Money) == 0:
                money.append(self.elements("order_type" ,2)[i].text)
                money.append(self.elements("order_zhushu" ,2)[i].text)
                money.append(random.randint(0 ,99))
            else:
                money.append(self.elements("order_type" ,2)[i].text)
                money.append(self.elements("order_zhushu" ,2)[i].text)
                money.append(int(max_Money))
            money_box[i].send_keys(str(money[3 + 3*i]))                                  
        return money

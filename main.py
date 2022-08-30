import logging
import datetime
import configparser
import os
import traceback
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from time import sleep
import pandas as pd
import configparser
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains


class MAPPING:

    def __init__(self):

        self.PATH = os.getcwd()

        # create logger
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.DEBUG)

        # create console handler
        ch = logging.StreamHandler()

        #create file handler 
        date = str(datetime.datetime.now().strftime('%d-%b-%Y %H_%M_%S %p'))

        fh = logging.FileHandler(f'{self.PATH}\\debug\\{date}.log',encoding='utf-8')

        # create formatter
        formatter = logging.Formatter('%(asctime)s - %(funcName)s - %(lineno)d - %(levelname)s - %(message)s',datefmt='%d/%b/%Y %I:%M:%S %p')

        # add formatter to ch
        ch.setFormatter(formatter)

        #add formatter to fh
        fh.setFormatter(formatter)

        # add ch to logger
        self.logger.addHandler(ch)

        #add fh to logger
        self.logger.addHandler(fh)


        #config.init file
        self.my_config_parser = configparser.ConfigParser()

        self.my_config_parser.read(f'{self.PATH}\\config\\config.ini')

        self.config = { 

        'email': self.my_config_parser.get('config','email'),
        'password': self.my_config_parser.get('config','password'),


        }
        
        self.err = {'NOT FOUND OPN':[],
                    'NOT FOUND WI':[]}
        
        
        
    def read_excel(self,file_name):
        
        self.map_dict = {'FA':{},
                         'Service': {},
                         'TEFR': {},
                         'RMK': {},
                         }

        self.df = pd.read_excel(f'{self.PATH}\\src\\{file_name}','Sheet2')
        
        self.logger.info(self.df)
    
        row_no = len(self.df.index)
        
        for row in range(row_no):

            opn_temp = self.df['OPN'][row].strip().split(', ')
            
            cert_temp = self.df['WI'][row].strip().split(', ')
            
            for opn in opn_temp:
                
                self.map_dict[self.df['JOBTYPE'][row]].update({f'{opn}':cert_temp})
                
        self.logger.info(' === MAPPING DICT ===')       
        self.logger.info(self.map_dict)
                     
                
    def add_mapping(self):
        
        self.driver=webdriver.Chrome()
        
        self.driver.get("https://fits/Cisco_FA/TrainingOperationMap/default.asp")

        for org in self.map_dict:
            
            self.logger.debug('ORG >>>  '+org)
                        
            for opn in self.map_dict[org]:
                
                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, f'//select[@id="select_job"]//option[ text() = "{org}" ]'))).click()
                
                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, f'//select[@name="select_model"]//option[text() = "*" ]'))).click()
                
                #select OPN
                self.logger.debug('OPN >>> '+opn+ ' >>> '+ str(self.map_dict[org][opn]))

                try:
                    
                    WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, f'//select[@name="select_opn"]//option[contains(text(),"{opn}")]'))).click()
                    
                except:
                    
                    self.err['NOT FOUND OPN'].append(opn)
                    
                    continue

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//input[@name="ok"]'))).click()
                
                
                #select wi
                
                for wi in self.map_dict[org][opn]:
                    
                    WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//select[@name="select_model2"]//option[text() = "All Models"]'))).click()
                    
                    wi_no = wi.split('_')[0]
                    
                    if 'VA' in wi_no:
                        
                        WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//select[@name="select_product2"]//option[text() = "Visual aid" ]'))).click()
 
                    else:
                        
                        WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//select[@name="select_product2"]//option[text() = "000_PE_All" ]'))).click()
                        
                    try:
                        
                        WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, F'//select[@name="select_opn2"]//option[contains(text(),"{wi_no}")]'))).click()
                      
                    except:
                        
                        self.err['NOT FOUND WI'].append(wi_no)
                        
                        continue
                        
                    WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//input[@name="add"]'))).click()
                
                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//input[@name="Submit"]'))).click()
                        
                    # handle with alert
                try:
                    WebDriverWait(self.driver, 3).until(ec.alert_is_present())

                    alert = self.driver.switch_to.alert
                    
                    alert.accept()
                    
                    self.logger.debug('alert accepted')

                except:
                    
                    self.logger.debug('no alert')
                        
                sleep(1)
                                    

    def delete_all_mapping(self):

        self.driver=webdriver.Chrome()
        
        self.driver.get("https://fits/Cisco_FA/TrainingOperationMap/default.asp")
        
        while True:

            try:
                
                del_opn = WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//tr[@bgcolor="Ivory"]//td')))[1].find_elements(By.TAG_NAME,'span')[0].text
                
                WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//tr[@bgcolor="Ivory"]//td')))[1].find_elements(By.TAG_NAME,'img')[0].click()

                self.logger.info('DELETE OPN >>> '+del_opn)
                # handle with 2 alert
                WebDriverWait(self.driver, 3).until(ec.alert_is_present())

                alert = self.driver.switch_to.alert
                
                alert.accept()
                
                alert.accept()
                
                self.logger.debug('alert accepted')

            except:
                
                traceback.print_exc()
                
                self.logger.debug('no alert')
                
                break
                
            sleep(1)
        

    def main(self):
        
        self.read_excel('FITS_MAPPING.xlsx')
        
        self.add_mapping()
        
        # self.delete_all_mapping()


if __name__ == '__main__':

    try:

        test = MAPPING()
        
        test.main()


    finally:
        
        test.logger.error(test.err)

        test.logger.critical("Traceback Error: "+traceback.format_exc())




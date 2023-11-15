import sys
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
import openpyxl,pandas
from selenium.webdriver.common.action_chains import ActionChains
import csv


chrome_options = Options()
chrome_options.add_argument("--headless")  # do not show the Chrome window
driver = webdriver.Chrome()
driver.maximize_window()


class LinkedinAutomation:
    
    @staticmethod
    def write_to_csv(data):
        csv_file_path = "./Variables/output.csv"
        
        fields = ["Sl. No", "Name", "Profile Link", "Title", "Company Name", "Status", "Type"]
        with open(csv_file_path, 'w', newline='') as csvfile:
            csv_writer = csv.DictWriter(csvfile, fieldnames=fields)
            csv_writer.writeheader()
            for idx, row in enumerate(data, start=1):
                row["Sl. No"] = idx
                csv_writer.writerow(row)

        print(f"Data has been written to {csv_file_path}")

    @staticmethod
    def read_ex():
        file_path="Variables\WEB SCRAPPING LINKDIN.xlsx"
        df=pandas.read_excel(file_path,engine='openpyxl')
        return df
        
    @staticmethod
    def web_automation_jsp(url, user, psw,title,location,industry):
        data = []

        try:

            driver.get(url)  # Navigate to the .jsp application

            # Step 1: find the username and write the username in the input
            WebDriverWait(driver, 60).until(
                EC.presence_of_all_elements_located((By.XPATH, '//*[@id="username"]')))

            username = driver.find_element(By.XPATH, '//*[@id="username"]')
            username.send_keys(user)

            # Step 2: find the password and write the password in the input
            
            password = driver.find_element(By.XPATH,'//*[@id="password"]')
            password.send_keys(psw)
            password.send_keys(Keys.RETURN)
            WebDriverWait(driver, 60).until(
                EC.presence_of_all_elements_located((By.XPATH, '//*[@id="global-nav-typeahead"]/input')))
            


            #Step 3: Search
            search=title+" "+industry+" "+location
            search_bar=driver.find_element(By.XPATH,'//*[@id="global-nav-typeahead"]/input')
            search_bar.send_keys(search)
            search_bar.send_keys(Keys.ENTER)
            WebDriverWait(driver, 60).until(
                EC.presence_of_all_elements_located((By.XPATH, '//*[@id="search-reusables__filters-bar"]/ul/li[1]/button')))
            see_all=driver.find_element(By.XPATH,'//*[@id="search-reusables__filters-bar"]/ul/li[1]/button').click()
            time.sleep(5)
            #Step 4: Select people
            print('entering for loop')
            page=1
            while(True):                    
                row=1
                while(True):
                    try:
                        if(row>7):
                            driver.find_element(By.TAG_NAME,'body').send_keys(Keys.PAGE_DOWN)
                            time.sleep(5)

                        print('clicking person')
                        print('check name')
                        try:
                            check_name=driver.find_element(By.XPATH,f'/html/body/div[5]/div[3]/div[2]/div/div[1]/main/div/div/div[2]/div/ul/li[{row}]/div/div/div/div[2]/div[1]/div[1]/div/span/span/a').text
                            print(check_name)
                        except Exception as e:
                            print(f"page {page} completed")
                            break
                        if(check_name=='LinkedIn Member'):
                            row+=1
                            continue
                        click_person=driver.find_element(By.XPATH,f'/html/body/div[5]/div[3]/div[2]/div/div[1]/main/div/div/div[2]/div/ul/li[{row}]').click()
                        
                        time.sleep(5)
                        print('type')
                        types=['Connect','Follow']
                        type=driver.find_element(By.XPATH,'/html/body/div[5]/div[3]/div/div/div[2]/div/div/main/section[1]/div[2]/div[3]/div/button[1]/span').text
                        if(type not in types):
                            type=driver.find_element(By.XPATH,'/html/body/div[5]/div[3]/div/div/div[2]/div/div/main/section[1]/div[2]/div[3]/div/button/span').text
                        print(type)
                        print('name')                       
                        name=driver.find_element(By.XPATH,'/html/body/div[5]/div[3]/div/div/div[2]/div/div/main/section[1]/div[2]/div[2]/div[1]/div[1]/span[1]/a/h1').text
                        driver.find_element(By.TAG_NAME,'body').send_keys(Keys.PAGE_DOWN)
                        time.sleep(5)
                        try:
                            print('title')
                            title=driver.find_element(By.XPATH,'/html/body/div[5]/div[3]/div/div/div[2]/div/div/main/section[3]/div[3]/ul/li[1]/div/div[2]/div/div[1]/div/div/div/div/span[1]').text
                            print('comapny')
                            company=driver.find_element(By.XPATH,'/html/body/div[5]/div[3]/div/div/div[2]/div/div/main/section[3]/div[3]/ul/li[1]/div/div[2]/div/div[1]/span[1]/span[1]').text
                        except Exception as e:
                            try:
                                print('title')
                                title=driver.find_element(By.XPATH,'/html/body/div[5]/div[3]/div/div/div[2]/div/div/main/section[4]/div[3]/ul/li[1]/div/div[2]/div/div[1]/div/div/div/div/span[1]').text
                                print('comapny')
                                company=driver.find_element(By.XPATH,'/html/body/div[5]/div[3]/div/div/div[2]/div/div/main/section[4]/div[3]/ul/li[1]/div/div[2]/div/div[1]/span[1]/span[1]').text
                            except Exception as e:
                                driver.find_element(By.TAG_NAME,'body').send_keys(Keys.PAGE_DOWN)
                                print('title')
                                title=driver.find_element(By.XPATH,'/html/body/div[5]/div[3]/div/div/div[2]/div/div/main/section[1]/div[2]/div[2]/div[1]/div[2]').text
                                print('comapny')
                                company=""

                        print('url')
                        link = driver.current_url
                        desc={"Name": name, "Profile Link": link, "Title": title, "Company Name": company, "Status": 0, "Type": type}
                        data.append(desc)
                        driver.back()
                        time.sleep(5)
                        # data collection
                        row+=1
                    except Exception as e:
                        print(e)
                        time.sleep(600)
                        break
                driver.find_element(By.TAG_NAME,'body').send_keys(Keys.PAGE_DOWN)
                time.sleep(3)
                try:
                    if(page==1):
                        next_page=driver.find_element(By.XPATH,'/html/body/div[5]/div[3]/div[2]/div/div[1]/main/div/div/div[5]/div/div/button[2]').click()
                        page=page+1
                    else:
                        next_page=driver.find_element(By.XPATH,'/html/body/div[5]/div[3]/div[2]/div/div[1]/main/div/div/div[2]/div/div[2]/div/button[2]').click()
                    time.sleep(5)                           
                except Exception as e:
                    print("Pages Completed/Writing Data.........")
                    LinkedinAutomation.write_to_csv(data)
        except Exception as e:
            print("An error occurred: ", e)

        finally:
            # Close the web driver
            print("Process Completed : ")
            driver.quit()


if __name__ == "__main__":
    df=LinkedinAutomation.read_ex()
    login_url = "https://www.linkedin.com/login"
    for _,rows in df.iterrows():
        LinkedinAutomation.web_automation_jsp(url=login_url, user=rows['Username '], psw=rows['Password'],title=rows['Tittle'],location=rows['LOCATION'],industry=rows['Industry'])
        time.sleep(600)





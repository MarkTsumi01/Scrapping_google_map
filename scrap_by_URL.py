import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd



df = pd.read_excel('Update.xlsx')
data = pd.read_excel('แก้ล่าสุด.xlsx')
processed_places = data['place'].tolist()

link_to_open = 'https://www.google.co.th/maps/'

option = webdriver.ChromeOptions()
option.add_argument("headless")

no_data = []
round = 0
same_names = 0
success = 0


for i in range(len(df)):
    # create variable to collect critilia place and point to loop
    place = df.iloc[i, 0]
    point = df.iloc[i, 1]
    points = int(point)
    url = df.iloc[i,3]
    driver = webdriver.Chrome()

    if place not in processed_places :

        driver.get(url)
        # maximize browser
        driver.maximize_window()
        # find search bar and enter place
        time.sleep(5)

        # search_input = driver.find_element(By.CLASS_NAME, 'searchboxinput')
        # search_input.send_keys(url)
        # time.sleep(5)
        # search_input.send_keys(Keys.ENTER)
        # time.sleep(5)

        check_review = True
        while check_review :
            try :
                review_tab = driver.find_element(By.XPATH,'//*[@id="QA0Szd"]/div/div/div[1]/div[3]/div/div[1]/div/div/div[2]/div[3]/div/div/button[2]/div[2]/div[2]')
                review_tab.click()
                time.sleep(5)
                # scroll_down = driver.find_element((By.XPATH,'//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div[1]')).click()
                scroll_down_area = driver.find_element(By.XPATH,'//*[@id="QA0Szd"]/div/div/div[1]/div[3]/div/div[1]/div/div/div[3]')
                time.sleep(10)
                for scroll in range(1, points):
                    scroll_down_area.send_keys(Keys.END)
                    time.sleep(0.2)
                check_review = False
                success += 1
            except Exception :
                print("Can't find review_tab")
                print("Can't find scroll tab")
                no_data.append(place)
                break
        time.sleep(5)


        # get current page to change it to html
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # find value with critiria and collect to variable
        review_all = soup.find_all('span',{'class':'wiI7pd'})
        name_all = soup.find_all('div',{'class':'d4r55'})
        star_all = soup.find_all('span',{'class':'kvMYJc'})
        list_aria_label = []
        for star_test in star_all:
            aria_label = star_test.get('aria-label')
            list_aria_label.append(aria_label)


        # debug total
        print(len(review_all))
        print(len(name_all))
        print(len(list_aria_label))
        print(place)


        # create empty list to collect
        place_name = []
        name = []
        review = []
        star = []
        source = []



        # Create loop to add value from soup to list
        for names in name_all:
            name.append(names.text)
        for reviews in review_all:
            review.append(reviews.text)
        for stars in list_aria_label:
            star.append(stars)
        for place_loop in range(len(name)):
            place_name.append(place)
        for source_loop in range(len(name)):
            source.append("google-map")


        # export list to excel
        all_data = pd.DataFrame([place_name,name,star,review,source])
        all_data = all_data.transpose()
        all_data.columns = ['place','name','star','comments','source']

        # all_data.to_excel(r'C:\Users\kanra\PycharmProjects\WebScrapping\Exportdata\Final.xlsx')

        data = pd.concat([data, all_data])
        data.to_excel('แก้ล่าสุด.xlsx')
        driver.quit()
    else :
        driver.quit()
        print(f"Same name = {place}")
        same_names +=1
    round += 1

len_data = len(no_data)
print(no_data)
print(f"Total round = : {round}")
print(f"Total no reviews = : {len_data}")
print(f"All same name = : {same_names}")
print(f"Success = {success}")

from bs4 import BeautifulSoup
import requests
from selenium import webdriver
import chromedriver_binary 
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook

endpoint = 'https://finance.yahoo.com/gainers'
def demo(count, offset):
   
    url = f"{endpoint}?count={str(count)}&offset={str(offset)}"
    print("========url=======", url)
    #=======================
    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.get(url)
    html = driver.execute_script('return document.body.innerHTML;')
    soup = BeautifulSoup(html,'lxml')
    #=======================
    '''
    page = requests.get(url, allow_redirects=True)
    soup = BeautifulSoup(page.content, "lxml")
    '''
    #=======================
    table = soup.find('table', attrs={'class':'W(100%)', 'data-reactid':'42'})
    table_body = table.find('tbody')
    rows = table_body.find_all('tr')
    #=======================
    wb = Workbook() # open new workbook
    ws = wb.create_sheet('Gainer')
    counter = 0
    headers = ['Symbol', 'Name', 'Price', 'Change', '%Change', 'Volume', 'Avg Vol','Market Cap','PE Ration','52 Week Range']
    ws.append(headers)  
    for row in rows:
        cols = row.find_all('td')
        print('=======>>>>', counter, '========')
        items = 0
        output = []
        for td in cols:
            value = ''
            if(items == 0):
                value = td.find('a').text 
            elif(items == 9):
                None
            else:
                value = td.text
            output.append(value)     
            items = items + 1   
        ws.append(output)        
        counter = counter + 1
    print('finished at ', counter)    
    if(counter > 0):
        offset = offset + 25
        print('new offset ' ,offset)
        #demo(count, offset)  
    print('saving.....')    
    wb.save('Gainers.xlsx')      


url = 'https://finance.yahoo.com/gainers'
count = 200
offset = 0
demo(count, offset)




from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import xlsxwriter

marka_list=[]
seri_list=[]
model_list=[]
options = Options()
options.headless = True
options.add_argument("--window-size=1920,1200")
driver = webdriver.Chrome()
for i in range(0,1000,20):
    driver.get("https://www.sahibinden.com/otomobil?pagingOffset="+str(i))
    for j in range(1,22,1):
        if j==4:
            continue
        if j==5:
            continue
        marka= driver.find_element_by_xpath("//*[@id='searchResultsTable']/tbody/tr["+str(j)+"]/td[2]")
        seri= driver.find_element_by_xpath("//*[@id='searchResultsTable']/tbody/tr["+str(j)+"]/td[3]")
        model= driver.find_element_by_xpath("//*[@id='searchResultsTable']/tbody/tr["+str(j)+"]/td[4]")
        marka_list.append(marka.text)
        seri_list.append(seri.text)
        model_list.append(model.text)


workbook = xlsxwriter.Workbook('write_list.xlsx')
worksheet = workbook.add_worksheet()

for row_num, data in enumerate(marka_list):
    worksheet.write(row_num, 0, data)

for row_num1, data in enumerate(seri_list):
    worksheet.write(row_num1, 1, data)

for row_num2, data in enumerate(model_list):
    worksheet.write(row_num2, 2, data)


workbook.close()
zipped_list = zip(marka_list,seri_list,model_list)
sorted_pairs = sorted(zipped_list)

tuples = zip(*sorted_pairs)
marka_list,seri_list,model_list = [list(tuple) for tuple in tuples]


workbook = xlsxwriter.Workbook('write_list1.xlsx')
worksheet = workbook.add_worksheet()

for row_num, data in enumerate(marka_list):
    worksheet.write(row_num, 0, data)

for row_num1, data in enumerate(seri_list):
    worksheet.write(row_num1, 1, data)

for row_num2, data in enumerate(model_list):
    worksheet.write(row_num2, 2, data)

workbook.close()

driver.quit()
import requests
import datetime
import xlsxwriter
from bs4 import BeautifulSoup
url = 'http://www.rakbankerd.com/price.php'
for year in range(2009, 2017):
    if (year % 4 == 0) and (year % 100 != 0 or year % 400 == 0):
        max_date = 366
    else:
        max_date = 365
     #print year
    for p_id in range(1, 78):
        print(p_id, year,)
        start_date = datetime.datetime(year, 1, 1)
        wb = xlsxwriter.Workbook('rice_%d_%d.xlsx'%(p_id,year))
        ws = wb.add_worksheet('%d'%year)
        r = 1
        for i in range(0, max_date):
            prms = dict()
            date = start_date + datetime.timedelta(i)
            prms['StartDate'] = date.strftime("%Y-%m-%d")
            prms['txtprice_search'] = ""
            prms['province'] = p_id
            req = requests.post(url,params={'ic': 1}, data=prms)
            soup = BeautifulSoup(req.text, 'html.parser')
            #print(date,)
            if u'ไม่พบข้อมูลค่ะ' in soup.get_text():
                continue
            else:
                for div in soup.find_all('div', {'class': 'panel panel-default'}):
                    head_div = div.find('div', {'class': 'panel-heading font-tahoma txt-18 text-center'})
                    body_div = div.find('tbody', {'class': 'txt-14'})
                    for tr in body_div.find_all('tr'):
                        ws.write(r, 0, head_div.get_text())
                        ws.write(r, 1, date.strftime("%Y-%m-%d"))
                        c = 2
                        for td in tr.find_all('td'):
                            ws.write(r, c, td.get_text())
                            c += 1
                        r += 1
        wb.close()


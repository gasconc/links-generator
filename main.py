import requests
import xlrd
import pandas as pd
import json
import progressbar

#progress bar
bar = progressbar.ProgressBar(maxval=100, \
    widgets=[progressbar.Bar('=', '[', ']'), ' ', progressbar.Percentage()])
bar.start()


#define Access_Token
AT=''
#Sheet creation
loc = ("Data.xlsx")
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0)

#initializations
tittles=[]
quantities=[]
unit_prices = []
external_references=[]
expiration_dates=[]
links=[]


for i in range(0,sheet.nrows):
    
    tittle= str(sheet.cell_value(i , 0))
    tittles.append(tittle)
    quantity= sheet.cell_value(i , 1)
    quantities.append(quantity)
    unit_price= sheet.cell_value(i , 2)
    unit_prices.append(unit_price)
    external_reference= str(sheet.cell_value(i , 3))
    external_references.append(external_reference)
    expiration_date= str(sheet.cell_value(i , 4))
    expiration_dates.append(expiration_date)
    body= { 
        "items":[{
            "title": tittle,
            "quantity": int(quantity),
            "unit_price":int(unit_price),
            "currency_id":"COP"
        }], 
        "expiration_date_to":expiration_date, 
        "external_reference":int(float(external_reference))
      
    }
    try: 
        result = requests.post('https://api.mercadopago.com/checkout/preferences?access_token='+AT, data= json.dumps(body))
    
        link= str(result.json()['init_point']) 
    
        links.append(link)
    except Exception:
        link = "error"
        links.append(link)
    bar.update((i/sheet.nrows) * 100)
    

bar.finish() 
raw_data={'tittle':tittles,'quantity':quantities, 'unit_price':unit_prices, 'expiration_date': expiration_dates, 'external_reference':external_references,'link':links}
df=pd.DataFrame(raw_data, columns=['tittle', 'quantity', 'unit_price','expiration_date','external_reference', 'link'])
output_workbook = 'output.xlsx'
df.to_excel(output_workbook, index=False)
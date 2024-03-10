# Libraries for Webscraping, Data Manipulation and Google Sheets + Cloud API

import requests
import re
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# U will need this later, to get access to the Google Sheets with the API key
# pip install gspread oauth2client pandas

#for single link testing (for API Connection testing) = "https://finance.yahoo.com/quote/HLAG.DE/key-statistics?p=HLAG.DE"

# Yahoo Finance Links to the Stocks we want to Scrape Data

link_list = ["https://finance.yahoo.com/quote/1U1.DE/key-statistics?p=1U1.DE", "https://finance.yahoo.com/quote/ADN1.DE/key-statistics?p=ADN1.DE", "https://finance.yahoo.com/quote/ADTN/key-statistics?p=ADTN", "https://finance.yahoo.com/quote/ADV.DE/key-statistics?p=ADV.DE", "https://finance.yahoo.com/quote/AAD.DE/key-statistics?p=AAD.DE", "https://finance.yahoo.com/quote/AOF.DE/key-statistics?p=AOF.DE", "https://finance.yahoo.com/quote/AG1.DE/key-statistics?p=AG1.DE", "https://finance.yahoo.com/quote/BYW.DE/key-statistics?p=BYW.DE", "https://finance.yahoo.com/quote/BFSA.DE/key-statistics?p=BFSA.DE", "https://finance.yahoo.com/quote/GBF.DE/key-statistics?p=GBF.DE", "https://finance.yahoo.com/quote/BVB.DE/key-statistics?p=BVB.DE", "https://finance.yahoo.com/quote/COK.DE/key-statistics?p=COK.DE", "https://finance.yahoo.com/quote/CEC.DE/key-statistics?p=CEC.DE", "https://finance.yahoo.com/quote/CWC.DE/key-statistics?p=CWC.DE", "https://finance.yahoo.com/quote/COP.DE/key-statistics?p=COP.DE", "https://finance.yahoo.com/quote/DMP.DE/key-statistics?p=DMP.DE", "https://finance.yahoo.com/quote/DBAN.DE/key-statistics?p=DBAN.DE", "https://finance.yahoo.com/quote/DWNI.DE/key-statistics?p=DWNI.DE", "https://finance.yahoo.com/quote/DEZ.DE/key-statistics?p=DEZ.DE", "https://finance.yahoo.com/quote/DRW3.DE/key-statistics?p=DRW3.DE", "https://finance.yahoo.com/quote/DUE.DE/key-statistics?p=DUE.DE", "https://finance.yahoo.com/quote/DWS.DE/key-statistics?p=DWS.DE", "https://finance.yahoo.com/quote/EUZ.DE/key-statistics?p=EUZ.DE", "https://finance.yahoo.com/quote/ELG.DE/key-statistics?p=ELG.DE", "https://finance.yahoo.com/quote/EKT.DE/key-statistics?p=EKT.DE", "https://finance.yahoo.com/quote/FIE.DE/key-statistics?p=FIE.DE", "https://finance.yahoo.com/quote/FTK.DE/key-statistics?p=FTK.DE", "https://finance.yahoo.com/quote/GFT.DE/key-statistics?p=GFT.DE", "https://finance.yahoo.com/quote/GYC.DE/key-statistics?p=GYC.DE", "https://finance.yahoo.com/quote/GLJ.DE/key-statistics?p=GLJ.DE", "https://finance.yahoo.com/quote/HABA.DE/key-statistics?p=HABA.DE", "https://finance.yahoo.com/quote/HDD.DE/key-statistics?p=HDD.DE", "https://finance.yahoo.com/quote/HBH.DE/key-statistics?p=HBH.DE", "https://finance.yahoo.com/quote/HYQ.DE/key-statistics?p=HYQ.DE", "https://finance.yahoo.com/quote/INH.DE/key-statistics?p=INH.DE", "https://finance.yahoo.com/quote/IOS.DE/key-statistics?p=IOS.DE", "https://finance.yahoo.com/quote/JST.DE/key-statistics?p=JST.DE", "https://finance.yahoo.com/quote/KCO.DE/key-statistics?p=KCO.DE", "https://finance.yahoo.com/quote/KTN.DE/key-statistics?p=KTN.DE", "https://finance.yahoo.com/quote/KSB3.DE/key-statistics?p=KSB3.DE", "https://finance.yahoo.com/quote/KWS.DE/key-statistics?p=KWS.DE", "https://finance.yahoo.com/quote/B4B.DE/key-statistics?p=B4B.DE", "https://finance.yahoo.com/quote/MOR.DE/key-statistics?p=MOR.DE", "https://finance.yahoo.com/quote/MUX.DE/key-statistics?p=MUX.DE", "https://finance.yahoo.com/quote/NA9.DE/key-statistics?p=NA9.DE", "https://finance.yahoo.com/quote/NOEJ.DE/key-statistics?p=NOEJ.DE", "https://finance.yahoo.com/quote/PAT.DE/key-statistics?p=PAT.DE", "https://finance.yahoo.com/quote/PBB.DE/key-statistics?p=PBB.DE", "https://finance.yahoo.com/quote/PFV.DE/key-statistics?p=PFV.DE", "https://finance.yahoo.com/quote/PNE3.DE/key-statistics?p=PNE3.DE", "https://finance.yahoo.com/quote/PSM.DE/key-statistics?p=PSM.DE", "https://finance.yahoo.com/quote/TPE.DE/key-statistics?p=TPE.DE", "https://finance.yahoo.com/quote/SFQ.DE/key-statistics?p=SFQ.DE", "https://finance.yahoo.com/quote/SZG.DE/key-statistics?p=SZG.DE", "https://finance.yahoo.com/quote/SHA.DE/key-statistics?p=SHA.DE", "https://finance.yahoo.com/quote/1SXP.DE/key-statistics?p=1SXP.DE", "https://finance.yahoo.com/quote/F3C.DE/key-statistics?p=F3C.DE", "https://finance.yahoo.com/quote/SGL.DE/key-statistics?p=SGL.DE", "https://finance.yahoo.com/quote/STO3.DE/key-statistics?p=STO3.DE", "https://finance.yahoo.com/quote/SBS.DE/key-statistics?p=SBS.DE", "https://finance.yahoo.com/quote/SZU.DE/key-statistics?p=SZU.DE", "https://finance.yahoo.com/quote/SMHN.DE/key-statistics?p=SMHN.DE", "https://finance.yahoo.com/quote/SYAB.DE/key-statistics?p=SYAB.DE", "https://finance.yahoo.com/quote/NCH2.DE/key-statistics?p=NCH2.DE", "https://finance.yahoo.com/quote/8TRA.DE/key-statistics?p=8TRA.DE", "https://finance.yahoo.com/quote/VAR1.DE/key-statistics?p=VAR1.DE", "https://finance.yahoo.com/quote/VBK.DE/key-statistics?p=VBK.DE", "https://finance.yahoo.com/quote/VOS.DE/key-statistics?p=VOS.DE", "https://finance.yahoo.com/quote/WAC.DE/key-statistics?p=WAC.DE", "https://finance.yahoo.com/quote/WUW.DE/key-statistics?p=WUW.DE"]

# define a valid Header for our Scraper, otherwise the Websites may block you, when you sent too much requests with an empty header

headers_http = {"User-Agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:99.0) Gecko/20100101 Firefox/99.0"}

# Hardcoded Column Names (after extracting them with the Code above)

headers_column = ['Market_Cap_(intraday)_', 'Enterprise_Value_', 'Trailing_P/E_', 'Forward_P/E_', 'PEG_Ratio_(5_yr_expected)_', 'Price/Sales_(ttm)', 
                  'Price/Book_(mrq)', 'Enterprise_Value/Revenue_', 'Enterprise_Value/EBITDA_', 'Beta_(5Y_Monthly)_', '52-Week_Change_3', 
                  'S&P500_52-Week_Change_3', '52_Week_High_3', '52_Week_Low_3', '50-Day_Moving_Average_3', '200-Day_Moving_Average_3', 'Avg_Vol_(3_month)_3', 
                  'Avg_Vol_(10_day)_3', 'Shares_Outstanding_5', 'Implied_Shares_Outstanding_6', 'Float_8', '%_Held_by_Insiders_1', '%_Held_by_Institutions_1', 
                  'Shares_Short__4', 'Short_Ratio__4', 'Short_%_of_Float__4', 'Short_%_of_Shares_Outstanding__4', 'Shares_Short_(prior_month_)_4', 
                  'Forward_Annual_Dividend_Rate_4', 'Forward_Annual_Dividend_Yield_4', 'Trailing_Annual_Dividend_Rate_3', 'Trailing_Annual_Dividend_Yield_3', 
                  '5_Year_Average_Dividend_Yield_4', 'Payout_Ratio_4', 'Dividend_Date_3', 'Ex-Dividend_Date_4', 'Last_Split_Factor_2', 'Last_Split_Date_3', 
                  'Fiscal_Year_Ends_', 'Most_Recent_Quarter_(mrq)', 'Profit_Margin_', 'Operating_Margin_(ttm)', 'Return_on_Assets_(ttm)', 'Return_on_Equity_(ttm)', 
                  'Revenue_(ttm)', 'Revenue_Per_Share_(ttm)', 'Quarterly_Revenue_Growth_(yoy)', 'Gross_Profit_(ttm)', 'EBITDA_', 'Net_Income_Avi_to_Common_(ttm)', 
                  'Diluted_EPS_(ttm)', 'Quarterly_Earnings_Growth_(yoy)', 'Total_Cash_(mrq)', 'Total_Cash_Per_Share_(mrq)', 'Total_Debt_(mrq)', 'Total_Debt/Equity_(mrq)', 
                  'Current_Ratio_(mrq)', 'Book_Value_Per_Share_(mrq)', 'Operating_Cash_Flow_(ttm)', 'Levered_Free_Cash_Flow_(ttm)']

# define the Lists, where we write the Webscraping Data in
data_list = []
finance_numbers = []
name_list = []

# Go over each Link in the List and send a request with our defined Header
# If the Response is 200 OK. Process with the Script otherwise print a Error Message
# with the soup Variable we get the Webiste content for each Link

for link in link_list:
    r = requests.get(link, headers=headers_http)
    if r.status_code == 200:
        soup = BeautifulSoup(r.content, "html.parser")
        
        # Get the beautiful Soup Object for all Finance Metrics with its HTMl Class
        finanz_zahlen = soup.find_all("td", class_="Fw(500) Ta(end) Pstart(10px) Miw(60px)", string=True)
        # Get the beautiful Soup Object for the Company Name of the Link
        firmenname = soup.find_all("h1", class_="D(ib) Fz(18px)", string=True)

        finance_numbers = []  # Initialize finance_numbers within the loop
        
        # Go over the finance Metrics in our Soup Object and append them to a list
        # Because every Value are Strings from the Soup Object
        # Yahoo Finance Gives Back Numbers like "24,3M" i convert Million, Billion etc. into floats

        for element in finanz_zahlen:
            finance_numbers.append(element.text)

        for idx, element in enumerate(finance_numbers):

            # Converts M into float Million and so on...
            if element.endswith("M"):
                finance_numbers[idx] = float(element[:-1]) * 1e6
            elif element.endswith("B"):
                finance_numbers[idx] = float(element[:-1]) * 1e9
            elif element.endswith("k"):
                finance_numbers[idx] = float(element[:-1]) * 1e3
            elif element.endswith("%"):
                finance_numbers[idx] = element

        data_list.extend(finance_numbers)  # Append finance_numbers to data_list after processing
        name_list.extend(firmenname)  # Append Company name to names list after processing

    else:
        print("Seems like there was an Error")

# Converting Finance Numbers into a Numpy Array and Reshape it into 60 columns (this sorts them as well)
data_array = np.array(data_list)
data_array = data_array.reshape(len(link_list), 60)
# Write the Array with our financial Metrics into the Dataframe
df_finance = pd.DataFrame(data_array, columns=headers_column)  

# do the same for the Company name
name_array = np.array(name_list)
name_array = name_array.reshape(len(link_list), 1)
df_names = pd.DataFrame(name_array, columns=["Company Name"])

# Get 1 Column for the Links, to look up further Data if Stock has good measurements

df_links = pd.DataFrame(link_list, columns=["Link_Yahoo_Finance"])

# create final DF to write into Google Sheets later
df_final = pd.concat([df_names, df_links, df_finance], axis=1)

# just a quick test if whole list was successfully crawled
print("Script Run Successfully!")

# Define the scope
scope = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']

# Insert your Google API JSON File here! Replace my old one (and add it to your Project Folder)
path = "evident-catcher-416200-7712669a9428.json" 

# Authenticate using credentials
credentials = ServiceAccountCredentials.from_json_keyfile_name(path, scope)
client = gspread.authorize(credentials)

# Open the Google Sheet Replace SDAX_DATA_DUMP with the Name of your Sheet
sheet = client.open('SDAX_Data_Dump').sheet1


# Write DataFrame to Google Sheet
# First, clear the existing contents of the sheet
sheet.clear()
# Then, write the DataFrame to the sheet
sheet.update([df_final.columns.values.tolist()] + df_final.values.tolist())
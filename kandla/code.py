import urllib.request
import os
from datetime import date, timedelta
import requests
import pandas as pd
import numpy as np


today = date.today().strftime("%d.%m.%Y")
previous_day = date.today() - timedelta(days=1)
previous_day = previous_day.strftime("%d.%m.%Y")
year_month = date.today().strftime("%Y/%m")

base_link = f'https://www.deendayalport.gov.in/wp-content/uploads/{year_month}/Daily-Berthing-List-'
ext = '.xlsx'

today_url = f"{base_link}{today}{ext}"
previous_day_url = f"{base_link}{previous_day}{ext}"

active_url = requests.get(today_url)

if active_url.status_code == 200:
    url = today_url
else:
    url = previous_day_url


def download_file_kandla(url, save_path, df_save_path):
    
    # historic_berth_data = pd.DataFrame(columns=["BERTH", "VESSEL NAME", "OPERATIONS", "CARGO", "CARGO QUANTITY", "UNIT", "ETB", "ETC", "AGENT", "REMARKS", "RECEIVED DATE"])
    # historic_berth_data.to_excel(os.path.join(df_save_path, "historic_berth.xlsx"), index=False)
    # print(historic_berth_data)
    
    historic_berth_data = pd.read_excel(df_save_path + "/historic_berth.xlsx", sheet_name=0)
    historic_arrived_data = pd.read_excel(df_save_path + "/historic_arrived.xlsx", sheet_name=0)

    try:

        os.makedirs(df_save_path, exist_ok=True)
        # Download the file
        urllib.request.urlretrieve(url,save_path)

        # Get the lineup received date
        df_temp = pd.read_excel(save_path, sheet_name=0, skiprows=2)
        df_temp = df_temp.iloc[:1]
        df_temp = df_temp.loc[:, df_temp.columns.str.contains('NEW')]
        df_temp.iloc[:, 0] = df_temp.iloc[:, 0].astype(str).str[-10:]
        lineup_received_date = df_temp.iloc[0,0]
        

        # --- Step 1: Extract Sheet1 and save as berth.xlsx ---

        df_berth = pd.read_excel(save_path, sheet_name=0, skiprows=5)
        df_berth = df_berth[[ 'PRIORITY', 'VCN No.', 'BERTH', 'VESSEL NAME', 'I/E', 'CARGO', 'QTY', 'UOM', 'COMM', 'ETC', 'AGENT', 'REMARKS']]
        df_berth['PRIORITY'] = df_berth['PRIORITY'].ffill()
        df_berth['BERTH'] = df_berth['BERTH'].astype(str)
        df_berth['BERTH'] = df_berth['BERTH'].str.replace(r"[()]", "", regex=True)
        df_berth['VESSEL NAME'] = df_berth['VESSEL NAME'].str.replace(r"[.]", "", regex=True)
        df_berth.loc[~df_berth['PRIORITY'].isin(['TT','OJ']), 'PRIORITY'] = np.nan
        df_berth['BERTH'] = np.where(df_berth['PRIORITY'].isna(),
                                      df_berth['BERTH'],
                                      df_berth['PRIORITY'].astype(str) + df_berth['BERTH'].astype(str)
                                      )
        df_berth = df_berth.drop(columns=['PRIORITY'])
        df_berth = df_berth.rename(columns={
            'VCN No.':  'VCN NO.',
            'I/E': 'OPERATIONS',
            'UOM': 'UNIT',
            'COMM': 'ETB',
            'QTY': 'CARGO QUANTITY'
        })
        df_berth    
        df_berth['RECEIVED DATE'] = lineup_received_date
        df_berth = df_berth[df_berth['BERTH'].notnull()]
        col_to_strip = [ 'BERTH', 'VESSEL NAME', 'OPERATIONS', 'CARGO', 'UNIT', 'AGENT', 'REMARKS']
        df_berth[col_to_strip]=df_berth[col_to_strip].apply(lambda x: x.str.strip())
        df_berth['ETB2'] = np.where(df_berth['REMARKS'] == "BERTHING TODAY",
                                    date.today(),
                                    df_berth['ETB'])
        df_berth = df_berth.drop(columns=['ETB'])
        df_berth = df_berth.rename(columns={'ETB2':'ETB'})

        df_berth.to_excel(os.path.join(df_save_path, "at_berth.xlsx"), index=False)

        historic_berth_data = pd.concat([historic_berth_data, df_berth], ignore_index=True)
        historic_berth_data["RECEIVED DATE"] = pd.to_datetime(historic_berth_data["RECEIVED DATE"], errors="coerce")
        historic_berth_data = historic_berth_data[historic_berth_data['RECEIVED DATE'] >= pd.to_datetime(date.today()-timedelta(days=1))]


        historic_berth_data.to_excel(os.path.join(df_save_path, "historic_berth.xlsx"), index=False)

        
        # --- Step 2: Extract Sheet2 and save as arrived.xlsx ---

        df_arrived = pd.read_excel(save_path, sheet_name=1, skiprows=1)

        df_arrived = df_arrived[['CJ/ OJ/ PPP', 'VCN No.', 'Vessel', 'Imp/ Exp', 'Cargo', 'Qty', 'UOM', 'Reporting', 'AGENT/STEV', 'REMARKS']]
        df_arrived['RECEIVED DATE'] = lineup_received_date
        df_arrived = df_arrived.rename(columns={
            'CJ/ OJ/ PPP': 'BERTH',
            'VCN No.':  'VCN NO.',
            'Vessel': 'VESSEL NAME',
            'Imp/ Exp': 'OPERATIONS',
            'Cargo': 'CARGO',
            'Qty': 'CARGO QUANTITY',
            'UOM': 'UNIT',
            'Reporting': 'ETA',
            'AGENT/STEV': 'AGENT',
            'REMARKS': 'REMARKS',
        })
        df_arrived = df_arrived[df_arrived['BERTH'].notnull()]
        df_arrived['VESSEL NAME'] = df_arrived['VESSEL NAME'].str.replace(r"[.]", "", regex=True)

        df_arrived.to_excel(os.path.join(df_save_path,"arrived.xlsx"), index=False)

        historic_arrived_data = pd.concat([historic_arrived_data, df_arrived], ignore_index=True)
        historic_arrived_data['RECEIVED DATE'] = pd.to_datetime(historic_arrived_data['RECEIVED DATE'], errors="coerce")
        historic_arrived_data = historic_arrived_data[historic_arrived_data['RECEIVED DATE'] >= pd.to_datetime(date.today()-timedelta(days=1))]
        historic_arrived_data.to_excel(os.path.join(df_save_path, "historic_arrived.xlsx"), index=False)

        # --- Step 3: Extract Sheet3 and save as expected.xlsx ---
        df_expected = pd.read_excel(save_path, sheet_name=2, skiprows=1)
        df_expected = df_expected[['CJ/ OJ/ PPP',  'VCN No.','Vessel', 'Imp/ Exp', 'Cargo', 'Qty', 'UOM', 'Estimated Arrival (Date & Time)', 'AGENT', 'Remarks']]
        df_expected['Received_Date'] = lineup_received_date
        df_expected = df_expected.rename(columns={
            'CJ/ OJ/ PPP':'BERTH',
            'VCN No.':  'VCN NO.',
            'Vessel':'VESSEL NAME',
            'Imp/ Exp':'OPERATIONS',
            'Cargo':'CARGO',
            'Qty':'CARGO QUANTITY',
            'UOM':'UNIT',
            'Estimated Arrival (Date & Time)':'ETA',
            'AGENT':'AGENT',
            'Remarks':'REMARKS',
            'Received_Date':'RECEIVED DATE'
        })
        df_expected['VESSEL NAME'] = df_expected['VESSEL NAME'].str.replace(r"[.]", "", regex=True)
        df_expected = df_expected[df_expected['BERTH'].notnull()]

        df_expected.to_excel(os.path.join(df_save_path,"expected.xlsx"), index=False)



    except Exception as e:

        print(e)
       
        return False
    
save_path = "./kandla/raw_lineup.xlsx"
df_save_path = "./kandla"

download_file_kandla(url, save_path, df_save_path)


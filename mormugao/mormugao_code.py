import urllib.request
import os
from datetime import date, timedelta
import requests
import pandas as pd
import numpy as np
import camelot

vessel_at_berth = 'https://mptgoa.gov.in/admin/portoperation-document/1776424265487664PORT%20POSITION%201.pdf'
vessel_arrived = 'https://mptgoa.gov.in/admin/portoperation-document/1776424285894858PORT%20POSITION%202.pdf'
vessel_expected = 'https://mptgoa.gov.in/admin/portoperation-document/1776424308318030PORT%20POSITION%203.pdf'

def lineup_mormugoa(vessel_at_berth, vessel_arrived, vessel_expected):

    try:
        urllib.request.urlretrieve(vessel_at_berth, save_path)

        tables = camelot.read_pdf(save_path, pages="all")

        if tables:
            with pd.ExcelWriter(df_save_path+"/raw_at_berth.xlsx", engine="openpyxl") as writer:
                for i, table in enumerate(tables):
                    table.df.to_excel(writer, sheet_name=f"Table_{i+1}", index=False)

        
        df_at_berth_1 = pd.read_excel(df_save_path+ "/raw_at_berth.xlsx", sheet_name="Table_1", skiprows=[1, 2])
        df_at_berth_1 = df_at_berth_1[[0, 1, 2, 3, 4, 10, 13, 14]]
        df_at_berth_1 = df_at_berth_1.rename(columns={
            0:'BERTH',
            1:'VESSEL NAME',
            2:'ARR/BER/NO',
            3:'CARGO/I/E',
            4:'AGENT/SHPR-RECVR/STEVHOOK/',
            10:'QTY I',
            13:'ETD',
            14:'REMARKS'         
        })
        df_at_berth_2 = pd.read_excel(df_save_path+ "/raw_at_berth.xlsx", sheet_name="Table_2", header=0)
        df_at_berth_2 = df_at_berth_2[[0, 1, 2, 3, 4, 10, 13, 14]]
        df_at_berth_2 = df_at_berth_2.rename(columns={
            0:'BERTH',
            1:'VESSEL NAME',
            2:'ARR/BER/NO',
            3:'CARGO/I/E',
            4:'AGENT/SHPR-RECVR/STEVHOOK/',
            10:'QTY I',
            13:'ETD',
            14:'REMARKS'         
        })

        df_at_berth_3 = pd.read_excel(df_save_path+ "/raw_at_berth.xlsx", sheet_name="Table_3",header=0)
        df_at_berth_3 = df_at_berth_3[[0, 1, 2, 3, 4, 10, 13, 14]]
        df_at_berth_3 = df_at_berth_3.rename(columns={
            0:'BERTH',
            1:'VESSEL NAME',
            2:'ARR/BER/NO',
            3:'CARGO/I/E',
            4:'AGENT/SHPR-RECVR/STEVHOOK/',
            10:'QTY I',
            13:'ETD',
            14:'REMARKS'         
        })
        
        Mormugoa_vessel_at_berth = pd.concat([df_at_berth_1, df_at_berth_2, df_at_berth_3], ignore_index=True)
        Mormugoa_vessel_at_berth = Mormugoa_vessel_at_berth[~Mormugoa_vessel_at_berth['BERTH'].str.contains("Working ", na=False)]
        Mormugoa_vessel_at_berth = Mormugoa_vessel_at_berth[~Mormugoa_vessel_at_berth['BERTH'].isin(["BERTH", ""])]
        Mormugoa_vessel_at_berth = Mormugoa_vessel_at_berth[Mormugoa_vessel_at_berth['BERTH'].notnull()]

        Mormugoa_vessel_at_berth[['VESSEL NAME', 'Drop']] = Mormugoa_vessel_at_berth['VESSEL NAME'].str.split('/', n=1, expand=True)
        Mormugoa_vessel_at_berth['ETA1'] = Mormugoa_vessel_at_berth['ARR/BER/NO'].str[0:11]
        Mormugoa_vessel_at_berth['ETB1'] = Mormugoa_vessel_at_berth['ARR/BER/NO'].str[12:23]
        Mormugoa_vessel_at_berth[['CARGO', 'OPERATIONS']] = Mormugoa_vessel_at_berth['CARGO/I/E'].str.split('/', n=1, expand=True)
        Mormugoa_vessel_at_berth[['AGENT', 'SHIPPER/RECEIVER']] = Mormugoa_vessel_at_berth['AGENT/SHPR-RECVR/STEVHOOK/'].str.split('/', expand=True)
        Mormugoa_vessel_at_berth = Mormugoa_vessel_at_berth[['BERTH', 'VESSEL NAME', 'ETA1', 'ETB1', 'ETD', 'OPERATIONS', 'CARGO', 'QTY I', 'AGENT', 'SHIPPER/RECEIVER', 'REMARKS']]
        
        Mormugoa_vessel_at_berth['VESSEL NAME'] = Mormugoa_vessel_at_berth['VESSEL NAME'].str.replace(r"[.]","", regex=True)
        Mormugoa_vessel_at_berth['ETA1'] = Mormugoa_vessel_at_berth['ETA1'].str.replace(r"[.]","-",regex=True)
        Mormugoa_vessel_at_berth['ETB1'] = Mormugoa_vessel_at_berth['ETB1'].str.replace(r"[.]","-",regex=True)


        today_month = int(date.today().strftime("%m"))
        today_year = int(date.today().strftime("%Y"))

        ETA_month = pd.to_numeric(
            Mormugoa_vessel_at_berth['ETA1'].str[3:4], errors="coerce"
        )

        Mormugoa_vessel_at_berth['ETA'] = np.where(
            Mormugoa_vessel_at_berth['VESSEL NAME'].isna(),
            "",
            np.where(
            (ETA_month > 11) & (today_month < 2),
            Mormugoa_vessel_at_berth['ETA1'].str[0:5]+"-"+str(today_year-1)+" "+Mormugoa_vessel_at_berth['ETA1'].str[6:11],
            Mormugoa_vessel_at_berth['ETA1'].str[0:5]+"-"+str(today_year)+" "+Mormugoa_vessel_at_berth['ETA1'].str[6:11]
            )
        )

        Mormugoa_vessel_at_berth['ETA'] = pd.to_datetime(Mormugoa_vessel_at_berth['ETA'], format='%d-%m-%Y %H:%M', errors="coerce")


        ETB_month = pd.to_numeric(
            Mormugoa_vessel_at_berth['ETB1'].str[4:5], errors="coerce"
        )

        Mormugoa_vessel_at_berth['ETB'] = np.where(
            (ETB_month > 11) & (today_month < 2),
            Mormugoa_vessel_at_berth['ETB1'].str[0:5]+"-"+str(today_year-1)+" "+Mormugoa_vessel_at_berth['ETB1'].str[6:11],
            Mormugoa_vessel_at_berth['ETB1'].str[0:5]+"-"+str(today_year)+" "+Mormugoa_vessel_at_berth['ETB1'].str[6:11]
            )

        Mormugoa_vessel_at_berth['ETB'] = pd.to_datetime(Mormugoa_vessel_at_berth['ETA'], format='%d-%m-%Y %H:%M', errors="coerce")

        Mormugoa_vessel_at_berth.to_excel(os.path.join(df_save_path,"test.xlsx"), index=False)


    except Exception as e:

        print(e)
       
        return False
    
save_path = "./Mormugao/raw_at_berth.pdf"
df_save_path = "./Mormugao"

lineup_mormugoa(vessel_at_berth, vessel_arrived, vessel_expected)
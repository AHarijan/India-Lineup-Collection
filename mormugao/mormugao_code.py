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

        
        df_at_berth = pd.read_excel(df_save_path+ "/raw_at_berth.xlsx", sheet_name=0, skiprows=3)

        df_at_berth.to_excel(os.path.join(df_save_path,"test.xlsx"), index=False)
        

    except Exception as e:

        print(e)
       
        return False
    
save_path = "./Mormugao/raw_at_berth.pdf"
df_save_path = "./Mormugao"

lineup_mormugoa(vessel_at_berth, vessel_arrived, vessel_expected)
# from datetime import date, timedelta
# import urllib.request
# import requests


# today = date.today().strftime("%d.%m.%Y")
# previous_day = date.today() - timedelta(days=1)
# previous_day = previous_day.strftime("%d.%m.%Y")



# base_link = 'https://www.deendayalport.gov.in/wp-content/uploads/2026/04/Daily-Berthing-List-'
# ext = '.xlsx'

# today_url = f"{base_link}{today}{ext}"
# previous_day_url = f"{base_link}{previous_day}{ext}"

# print(today_url)
# print(previous_day_url)

# headers = {
#     "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
# }

# active_url = requests.get(today_url, headers=headers)

# if active_url.status_code == 200:
#     url = today_url
# else:
#     url = previous_day_url

# print(url)

# headers = {
#     "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
# }
# save_path = "../raw.xlsx"
# active_url = requests.get(today_url, headers=headers)

# if active_url.status_code == 200:
#     url = today_url
# else:
#     url = previous_day_url

# # ✅ download whichever url was selected
# response = requests.get(url, headers=headers)
# with open(save_path, 'wb') as f:
#     f.write(response.content)

# print(f"Downloaded: {url}")
df_save_path = "./Mormugao"
print(df_save_path+"/at_berth.xlsx")
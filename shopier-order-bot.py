import pandas as pd
import requests
from datetime import datetime
from dotenv import load_dotenv
import os

load_dotenv()

url = "https://api.shopier.com/v1/orders"
date_start = "2022-07-21"
time_start = "13:24:51+0300"
date_end = "2023-04-05"
time = datetime.now().strftime("%Y-%m-%dT%H:%M:%S+0300")

params = {
    "dateStart": f"{date_start}T{time_start}",
    "dateEnd": f"{time}",
    "fulfillmentStatus": "unfulfilled",
    "limit": 50,
    "page": 1,
    "sort": "dateDesc"
}

headers = {
    "accept": "application/json",
    "authorization": os.getenv("API_KEY")
}
response = requests.get(url, params=params, headers=headers)
if response.status_code == 200:
    orders = response.json()
    unfulfilled_orders = []
    for order in orders:
        if order["status"] == "unfulfilled":
            unfulfilled_orders.append(order)
    print(f"Found {len(unfulfilled_orders)} unfulfilled orders.")
    
    try:
        df = pd.read_excel("unfulfilled_orders.xlsx")
    except FileNotFoundError:
        df = pd.DataFrame()

    new_data = []
    for order in unfulfilled_orders:
        name = order['shippingInfo']['firstName'] + ' ' + order['shippingInfo']['lastName']
        email = order['shippingInfo']['email']
        phone = order['shippingInfo']['phone']
        price = order['totals']['total']
        date_created = order['dateCreated']
        new_data.append([name, email, phone, price, date_created])
    new_df = pd.DataFrame(new_data, columns=["Name", "Email", "Phone", "Price", "Date Created"])
    df = pd.concat([df, new_df], ignore_index=True)


    df.to_excel("unfulfilled_orders.xlsx", index=False)

else:
    print(f"Error: {response.status_code} - {response.text}")
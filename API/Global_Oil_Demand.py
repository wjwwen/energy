import requests
import json
import pandas as pd
import urllib.parse
from pprint import pprint

username="USERNAME"
password="PASSWORD"
api_key="APIKEY"

url = "https://api.platts.com/mgodf/v1/demand"

def get_token(username, password, apikey):
  body = {
    "username": username,
    "password": password
  }
  headers = {
    "appkey": apikey
  }
  try:
    r = requests.post("https://api.platts.com/auth/api", data=body, headers=headers)
    r.raise_for_status()
    return r.json()["access_token"]
  except Exception as err:
    if r.status_code >= 500:
      print(err)
    else:
      print(r.status_code, r.json())

def get_history_assessments(username, password, apikey):
  token = get_token(username, password, apikey)

  # quotes are required around each symbol
  #symbols = [ '"' + x + '"' for x in symbols]
  #params = {
  #  "filter": f"symbol in ({','.join(symbols)})"
  #}
  headers = {
    "Authorization": f"Bearer {token}",
    "appkey": apikey
  }

  try:
    #print (params)
    r = requests.get(url,  headers=headers)
    r.raise_for_status()
    return r.json()
  except Exception as err:
    if r.status_code >= 500:
      print (err)
    else:
      print(r.status_code, r.json())
    raise

data = get_history_assessments(username, password, api_key)
#pprint(data["results"])

#json_str = json.loads(response.text)

# Writing to sample.csv
df = pd.json_normalize(data['results'], meta="year")
df.to_csv('sample.csv')

#print (pivot)
print("File saved with resultset in sample.csv...")
import json
import requests
import pandas as pd
from pandas import json_normalize

username = ""
password = ""

# https://api.connect.ihsmarkit.com/swagger/ui/index#/

query = "https://api.connect.ihsmarkit.com/dataplatform/v1/odata/Oil_Markets_Midstream_And_Downstream_API_Data_Simplified"

while query:
    response = requests.get(query, auth=(username, password))
    response.raise_for_status()

    response_content = json.loads(response.content)
    print(response_content)
    
    query = response_content.get("@odata.nextLink", None)
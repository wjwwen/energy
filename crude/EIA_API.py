import pandas as pd
import requests
# documentation: https://www.eia.gov/opendata/documentation.php

# US EIA API Key
api_key = ''

# PADD Names to Label Columns
# PADD 1 - East Coast
# PADD 2 - Midwest
# PADD 3 - Gulf Coast
# PADD 4 - Rocky Mountain
# PADD 5 - West Coast, AK, HI
PADD = ['PADD 1','PADD 2','PADD 3','PADD 4','PADD 5']

# Crude consumption by PADD
PADD_Key = ['PET.MCRRIP12.M',
'PET.MCRRIP22.M',
'PET.MCRRIP32.M',
'PET.MCRRIP42.M',
'PET.MCRRIP52.M']

final_data = []

# Choose start and end dates
startDate = '2020-01-01'
endDate = '2021-01-01'

# Pull in data via EIA API
for i in range(len(PADD_Key)):
    url = 'https://api.eia.gov/series/?api_key=' + api_key + '&series_id=' + PADD_Key[i]
    r = requests.get(url)
    json_data = r.json()
    
    if r.status_code == 200:
        print('Success')
    else:
        print('Error')
    
    df = pd.DataFrame(json_data.get('series')[0].get('data'),
                      columns = ['Date', PADD[i]])
    df.set_index('Date', drop=True, inplace=True)
    final_data.append(df)
import requests
from bs4 import BeautifulSoup
import re

query = "russia ukraine"
search = query.replace(' ', '+')
results = 500
url = (f"https://www.google.com/search?tbm=nws&q=oil") 

requests_results = requests.get(url)
soup_link = BeautifulSoup(requests_results.content, "html.parser")
links = soup_link.find_all("a")

for link in links:
    link_href = link.get('href')
    if "url?q=" in link_href and not "webcache" in link_href:
      title = link.find_all('h3')
      if len(title) > 0:
          # get link 
          # print(link.get('href').split("?q=")[1].split("&sa=U")[0])
          print(title[0].getText())
          titles = title[0].getText()
          print("------")
          
# Need to expand with writing into CSV file/email
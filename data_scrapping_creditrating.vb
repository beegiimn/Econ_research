pip install openpyxld

import requests
from bs4 import BeautifulSoup
import pandas as pd

# specify the URL of the webpage containing the links
url = 'http://www.worldgovernmentbonds.com/world-credit-ratings/'

# send a GET request to the webpage and store the response
response = requests.get(url)

# parse the HTML content of the response
soup = BeautifulSoup(response.content, 'html.parser')

# find all the links on the page
links = soup.find_all('a')

# iterate over the links
for link in links:
    # get the href attribute of the link
    link_url = link.get('href')
    # check if the link is not empty and starts with "http"
    if link_url and link_url.startswith('http'):
        # send a GET request to the link's URL
        link_response = requests.get(link_url)
        # parse the HTML content of the response
        link_soup = BeautifulSoup(link_response.content, 'html.parser')
        # find the last table on the page
        tables = link_soup.find_all('table')
        if tables:
            table = tables[-1]
            # extract the data from the table
            data = []
            for row in table.find_all('tr'):
                cols = row.find_all('td')
                cols = [col.text for col in cols]
                data.append(cols)
            # convert the data to a DataFrame
            df = pd.DataFrame(data)
            # save the data to an Excel file
            df.to_excel(f"{link_url.split('/')[-1]}.xlsx")
        else:
            print(f"No table found on {link_url}")

        # extract the data from the table
        data = []
        for row in table.find_all('tr'):
            cols = row.find_all('td')
            cols = [col.text for col in cols]
            data.append(cols)
        # convert the data to a DataFrame
        df = pd.DataFrame(data)
        # save the data to an Excel file
        if link_url.split('/')[-1]:
            df.to_excel(f"{link_url.split('/')[-1]}.xlsx")
        else:
            df.to_excel("default_file_name.xlsx")


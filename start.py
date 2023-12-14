import sys
import io
import requests 
import random
import sys
from bs4 import BeautifulSoup
import pandas as pd

# Change the console encoding to 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

menURL = "https://www.ncaa.com/rankings/basketball-men/d1/associated-press"
womenURL ="https://www.ncaa.com/rankings/basketball-women/d1/associated-press"

if random.choice([True, False]):  # Randomly choose between men and women URLs
    page = requests.get(menURL)
else:
    page = requests.get(womenURL)

soup = BeautifulSoup(page.content, "html.parser")
rankings_table = soup.find(class_="sticky")

# Check if the table is found
if rankings_table:
    # Extracting rows from the table body
    rows = rankings_table.find('tbody').find_all('tr')

    # Initialize empty lists to store data
    ranks = []
    teams = []
    records = []
    points = []
    previous_ranks = []

    # Loop through each row
    for row in rows:
        # Extracting data from each column (td)
        columns = row.find_all('td')
        rank = columns[0].text.strip()
        team = columns[1].text.strip()
        record = columns[2].text.strip()
        point = columns[3].text.strip()
        previous_rank = columns[4].text.strip()

        # Append data to lists
        ranks.append(rank)
        teams.append(team)
        records.append(record)
        points.append(point)
        previous_ranks.append(previous_rank)

    # Create a DataFrame
    df = pd.DataFrame({
        'Rank': ranks,
        'Team': teams,
        'Record': records,
        'Points': points,
        'Previous Rank': previous_ranks
    })

    # Save DataFrame to Excel file
    if menURL:
        df.to_excel('NCAA_MEN_RANKINGS.xlsx', index=False)
        print("Excel file Men updated successfully.")
    if womenURL:
        df.to_excel('NCAA_WOMEN_RANKINGS.xlsx', index=False)
        print("Excel file Women updated successfully.")
    else: 
        print('Error: Rankings Table not found on the webpage')
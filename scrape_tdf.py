import os
from urllib.request import urlopen, Request
from bs4 import BeautifulSoup
import pandas as pd

# Create an Excel writer object
file_path = os.path.join(os.getcwd(), "tour_de_france_stages2.xlsx")
with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
    # Loop through stages 4 to 16
    for stage_number in range(19, 22):
        url = f"https://www.procyclingstats.com/race/tour-de-france/2024/stage-{stage_number}"
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
        request = Request(url, headers=headers)

        html = urlopen(request)
        soup = BeautifulSoup(html, features="lxml")

        table = soup.find('table')

        table_headers = [th.getText().strip() for th in table.findAll('th')]
        rows = table.findAll('tr')[1:]
        rows_data = [[td.getText().strip() for td in row.findAll('td')] for row in rows]

        stage_data = pd.DataFrame(rows_data, columns=table_headers)

        # Save DataFrame to a new sheet in the Excel file
        sheet_name = f"Stage_{stage_number}"
        stage_data.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"Sheet '{sheet_name}' has been created and saved in the Excel file.")

print(f"Excel file '{file_path}' has been created and saved locally.")

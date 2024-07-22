import os
import re
from datetime import datetime
from typing import List

from pyppeteer import launch

from processor.processor import Processor
from utils.utils import TEAMS_URL, DEFAULT_URL, safe_int, safe_float

"""
This class is the processor to scrape the data from the njcaa website 
relative with baseball.
"""
class Scraper:
    processor: Processor
    def __init__(self):
        self.processor = Processor()
    """
    This method is the task to scrape the data
    """
    async def scrapping_task(self) -> None:
        print("Starting the Program...")
        print("\n")
        link = await self.scrape_real_url_data(TEAMS_URL)
        # Get the current datetime
        now = datetime.now()
        # Format the date and time in the specified format: year_month_date_hour_minute_second
        formatted_now = now.strftime("%Y_%m_%d_%H_%M_%S")

        for realurl in link:
            # Use a regular expression to find the required text
            match = re.search(r'/teams/([^/?]+)', realurl)
            print("Scraping " + match.group(1) + " data...")
            para_url = DEFAULT_URL + realurl
            hitters_data, pitchers_data = await self.scrape_table_data(para_url)

            if match:
                extracted_text = "reports/" + formatted_now + "/" + match.group(1) + ".xlsx"
            else:
                extracted_text = "reports/" + formatted_now + "/error.xlsx"

            if not os.path.exists("reports/" + formatted_now):
                # If it doesn't exist, create it
                os.makedirs("reports/" + formatted_now)

            # Now specify headers directly before calling save_table_as_excel
            hitters_headers = ['NO.', 'NAME', 'YR.', 'POS', 'BB', 'K', 'AVG', 'OBP', 'SLG', 'GO/FO', 'PA']
            pitchers_headers = ['NO.', 'NAME', 'YR.', 'POS', 'IP', 'BB', 'K/9', 'ERA', 'WHIP']

            # Save the hitters data
            self.processor.save_table_as_excel(hitters_headers, hitters_data, extracted_text, 'Hitters', append_mode=False)
            # Save the pitchers data in the same file, new sheet
            self.processor.save_table_as_excel(pitchers_headers, pitchers_data, extracted_text, 'Pitchers', append_mode=True)
            print("Outputed " + match.group(1) + ".excel file!")
            print("\n")

        print("The end!!!")

    """
    This method is to scrape the teams url from the base one. 
    """
    async def scrape_real_url_data(self, url: str) -> List[str]:

        # Launch the browser
        browser = await launch({
            'executablePath': "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
            'headless': True
        })
        page = await browser.newPage()

        await page.goto(url, {'timeout': 90000})

        # # Wait for at least one element with the class '.content-col' to be loaded
        await page.waitForSelector('.content-col')

        elements = await page.querySelectorAll('.content-col')

        revision_20_elements = await elements[1].querySelectorAll("ul")

        li_elements = await revision_20_elements[13].querySelectorAll("li")

        stats_links = []

        for li_element in li_elements:
            # Extract the text and href for each 'a' tag within the 'li' element
            a_tags = await li_element.querySelectorAll('a')
            for a_tag in a_tags:
                link_text = await page.evaluate('(element) => element.textContent', a_tag)
                link_href = await page.evaluate('(element) => element.getAttribute("href")', a_tag)
                if 'Stats' in link_text:
                    stats_links.append(link_href)

        # Close the browser
        await browser.close()

        return stats_links

    """
    This method is to scrape the real data from the url. 
    """
    async def scrape_table_data(self, url: str) -> (List[any], List[any]):

        hitters_data = []
        pitchers_data = []

        browser = await launch({
            'executablePath': "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
            'headless': True,  # Change to True if you don't need the browser to open up
        })
        page = await browser.newPage()

        await page.goto(url, {'timeout': 90000})

        # Wait for the table to load
        await page.waitForSelector('table')

        table_elements = await page.querySelectorAll('table')

        first_table = await table_elements[2].querySelectorAll("tr")

        for list in first_table:
            sub_data = []
            temp = await page.evaluate('(element) => element.textContent', list)
            temp = temp.replace(" ", "")
            temp = temp.replace("\n", ",")
            temp = temp.replace(",,,", ",")
            temp = temp.replace(",,", ",")
            parts = temp.split(",")
            if len(parts) < 21 and len(parts) > 4:
                parts.insert(4, ' ')
                temp = ",".join(parts)

            if parts[1] != '\xa0':
                sub_data.append(safe_int(parts[1]))  # NO.
                sub_data.append(parts[2])  # NAME
                sub_data.append(parts[3])  # YR.
                sub_data.append(parts[4])  # POS
                sub_data.append(safe_int(parts[13]))  # BB
                sub_data.append(safe_int(parts[14]))  # K
                sub_data.append(safe_float(parts[17]))  # AVG
                sub_data.append(safe_float(parts[18]))  # OBP
                sub_data.append(safe_float(parts[19]))  # SLG

            hitters_data.append(sub_data)

        second_table = await table_elements[4].querySelectorAll("tr")

        for list in second_table:
            temp = await page.evaluate('(element) => element.textContent', list)
            temp = temp.replace(" ", "")
            temp = temp.replace("\n", ",")
            temp = temp.replace(",,,", ",")
            temp = temp.replace(",,", ",")

            parts = temp.split(",")
            if len(parts) < 17 and len(parts) > 4:
                parts.insert(4, ' ')
                temp = ",".join(parts)

            if parts[1] != '\xa0':
                for x in hitters_data:
                    if len(x) > 1 and x[1] == parts[2]:
                        x.append(safe_float(parts[14]))  # GO/FO
                        x.append(safe_int(parts[15]))  # PA
                        break

        third_table = await table_elements[6].querySelectorAll("tr")

        for list in third_table:
            sub_data = []
            temp = await page.evaluate('(element) => element.textContent', list)
            temp = temp.replace(" ", "")
            temp = temp.replace("\n", ",")
            temp = temp.replace(",,,", ",")
            temp = temp.replace(",,", ",")

            parts = temp.split(",")
            if len(parts) < 22 and len(parts) > 4:
                parts.insert(4, ' ')
                temp = ",".join(parts)

            if parts[1] != '\xa0':
                sub_data.append(safe_int(parts[1]))  # NO.
                sub_data.append(parts[2])  # NAME
                sub_data.append(parts[3])  # YR.
                sub_data.append(parts[4])  # POS
                sub_data.append(safe_float(parts[11]))  # IP
                sub_data.append(safe_int(parts[15]))  # BB
                sub_data.append(safe_float(parts[17]))  # K/9
                sub_data.append(safe_float(parts[19]))  # ERA
                sub_data.append(safe_float(parts[20]))  # WMIP

            pitchers_data.append(sub_data)

        hitters_data = hitters_data[1:]
        pitchers_data = pitchers_data[1:]

        await browser.close()

        # Return the Data
        return hitters_data, pitchers_data
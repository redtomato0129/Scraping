import asyncio

from scraper.scraper import Scraper

scraper = Scraper()
asyncio.run(scraper.scrapping_task())
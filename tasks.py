import os
import time
import re
import logging
import requests
import shutil
from datetime import datetime, timedelta
from pathlib import Path
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from robocorp.tasks import task
from robocorp import workitems, browser, vault

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class DateCalculator:
    """This class contains the logic to calculate the time period input to search news between that days."""
    
    @staticmethod
    def calculate(actual_date, months):
        """Calculates the start date by subtracting the given number of months."""
        logging.info("Calculating date from %d months ago.", months)
        start_date = actual_date.replace(day=1)
        for _ in range(months):
            start_date -= timedelta(days=start_date.day)
        return start_date, actual_date


class ExcelCreator:
    """Class to handle the creation and management of an Excel file."""
    
    def __init__(self, filename):
        self.filename = filename
        self.workbook = Workbook()
        self.sheet = self.workbook.active

    def create_headers(self, headers):
        """Create headers for the Excel sheet."""
        logging.info("Creating Excel headers: %s", headers)
        self.sheet.append(headers)

    def add_row(self, data):
        """Add a row of data to the Excel sheet."""
        logging.info("Adding row to Excel: %s", data)
        self.sheet.append(data)

    def save_file(self):
        """Save the Excel file to the specified directory."""
        output_dir = Path(os.environ.get('ROBOT_ARTIFACTS'))
        output_path = output_dir / self.filename
        logging.info("Saving Excel file to %s", output_path)
        self.workbook.save(filename=output_path)


class FoxNewsSearch:
    """Class to handle web scraping and search operations on Fox News."""

    def __init__(self):
        logging.info("Initializing Selenium WebDriver.")
        options = Options()
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--headless")
        options.add_argument("--remote-debugging-port=9222")

        self.driver = webdriver.Chrome(options=options)
        self.wait = WebDriverWait(self.driver, 10)

    def click(self, xpath):
        """Click an element on the page using XPath with the rpa library."""
        logging.debug("Clicking element with XPath: %s", xpath)
        time.sleep(2)
        self.driver.find_element(By.XPATH, xpath).click()

    def search(self, keyword):
        """Perform a search on Fox News. search the button by XPATH, then click the button then write the prhase to search
         then enter to search"""
        logging.info("Performing search for: %s", keyword)
        button_xpath = "//div[@class='search-toggle tablet-desktop']/a[@class='js-focus-search']" #Search Button
        search_input_xpath = "//input[@type='text' and (@aria-label='search foxnews.com' or @placeholder='Search foxnews.com') and @name='q']" #Search text input
        self.click(button_xpath)
        self.driver.find_element(By.XPATH, search_input_xpath).send_keys(keyword)
        self.driver.find_element(By.XPATH, search_input_xpath).send_keys(Keys.ENTER)

    @staticmethod
    def download_image(image_url, filename):
        """Download an image from the specified URL. Review if exists using the status code from Url
        and then download the image to the robot folder"""
        output_dir = Path(os.environ.get('ROBOT_ARTIFACTS'))
        output_path = output_dir / filename
        logging.info("Downloading image from %s to %s", image_url, output_path)
        response = requests.get(image_url, stream=True)
        if response.status_code == 200: 
            with open(output_path, 'wb') as f:
                shutil.copyfileobj(response.raw, f)
            logging.info("Image downloaded successfully.")
        else:
            logging.error("Error downloading image.")

    @staticmethod
    def phrase_counter(text, phrase):
        """Count occurrences of a phrase in the given text by spliting the word and the phrase to compare char by char if 
        the word is in the phrase"""
        logging.debug("Counting occurrences of phrase '%s' in text.", phrase)
        text_words = text.lower().split()
        phrase_words = phrase.lower().split()
        return sum(
            1 for i in range(len(text_words) - len(phrase_words) + 1)
            if text_words[i:i + len(phrase_words)] == phrase_words
        )

    @staticmethod
    def contains_money(text):
        """Check if the text contains any monetary references by using a Regex."""
        logging.debug("Checking for money references in text.")
        money_pattern = r'\$[\d,]+(?:\.\d+)?|\b\d+\s*(?:dollars|USD)\b'
        return bool(re.findall(money_pattern, text))

    def close(self):
        """Close the Selenium WebDriver."""
        logging.info("Closing WebDriver.")
        time.sleep(10)
        self.driver.quit()


@task
def minimal_task():
    logging.info("Starting minimal task.")
    
    # Retrieve work item payload
    item = workitems.inputs.current
    logging.info("Received payload: %s", item.payload)
    
    payload = item.payload
    # Set up of the variables to use, according to setups in robocorp
    date_parameter = payload.get('Month', '0')
    phrase_to_search = payload.get('Phrase', 'Economy')
    category_to_search = payload.get('Category', 'LatinAmerica')
    phrase_category_search = f"{phrase_to_search} in {category_to_search}"
    actual_date = datetime.now()
    if date_parameter < 0:
        logging.error("Invalid date parameter: %d", date_parameter)
        return

    # Calculate the date range
    date_parameter -= 1
    start_date, actual_date = DateCalculator.calculate(actual_date, date_parameter)
    past_month, past_year = start_date.month, start_date.year
    current_month, current_day, current_year = actual_date.month, actual_date.day, actual_date.year
    year_index = (current_year - past_year) + 1

    # Convert int variables to 00 String format
    past_month_formatted = f"0{past_month}" if past_month < 10 else str(past_month)
    current_month_formatted =  f"0{current_month}" if current_month < 10 else str(current_month)
    current_day_formatted =  f"0{current_day}" if current_day < 10 else str(current_day)
    year_index_formatted =  f"0{year_index}" if year_index < 10 else str(year_index)

    # Init chrome
    logging.info("Configuring browser for automation.")
    browser.configure(
        browser_engine="chromium",
        screenshot="only-on-failure",
        headless=False,
    )

    # Retrieve secrets for authentication
    secrets = vault.get_secret('Rpa_Challenge')
    page = browser.goto(secrets['url'])
    time.sleep(3)

    # Perform the search on the website
    logging.info("Performing search on the website.")
    page.click("//div[@class='search-toggle tablet-desktop']/a[@class='js-focus-search']")
    page.fill("//input[@type='text' and (@aria-label='search foxnews.com' or @placeholder='Search foxnews.com') and @name='q']", phrase_category_search)
    page.click("//input[@type='submit' and @aria-label='submit search' and @class='resp_site_submit']")
    time.sleep(5)

    """ Select Date Range (From and To)
    As we have similar selectors for Month/Day selector, we rather to use the full XPath selector for this case. 
    """
    logging.info("Selecting date range for the search.")
    page.click("//*[@id='wrapper']/div[2]/div[1]/div/div[2]/div[3]/div[1]/div[1]") 
    page.click(f"//li[@id='{past_month_formatted}' and @class='{past_month_formatted}' and .='{past_month_formatted}']")
    page.click("//*[@id='wrapper']/div[2]/div[1]/div/div[2]/div[3]/div[1]/div[2]/button")
    page.click("//li[@id='01' and @class='01' and .='01']")
    page.click("//*[@id='wrapper']/div[2]/div[1]/div/div[2]/div[3]/div[1]/div[3]/button")
    page.click("//li[@id='2024']")
    time.sleep(1)
    page.click("//*[@id='wrapper']/div[2]/div[1]/div/div[2]/div[3]/div[2]/div[1]/button")
    page.click(f"//li[@id='{current_month_formatted}' and @class='{current_month_formatted}' and .='{current_month_formatted}']")
    page.click("//*[@id='wrapper']/div[2]/div[1]/div/div[2]/div[3]/div[2]/div[2]/button")
    page.click(f"//li[@id='{current_day_formatted}' and @class='{current_day_formatted}' and .='{current_day_formatted}']")
    page.click("//*[@id='wrapper']/div[2]/div[1]/div/div[2]/div[3]/div[2]/div[3]/button")
    page.click(f"//li[@id='{year_index_formatted}' and @class='{year_index_formatted}' and .='{year_index_formatted}']")

    # Start searching
    logging.info("Starting search.")
    page.click("//div[@class='button']/a[text()='Search']")
    time.sleep(3)
    
    logging.info("Entering the loop to load more news.")
    while not page.is_hidden("//span[text()='Load More']"):
        page.click("//span[text()='Load More']")
        time.sleep(3)

    logging.info("Finished loading news.")

    """ Get the number of news articles
        As we need this is a dynamic value in the page, we rather to use the full Xpath to retrieve the value in the span

    """
    news_amount = int(page.inner_text("//div[@class='num-found']/span[2]/span"))
    logging.info("Amount of news articles found: %d", news_amount)

    # Create Excel table
    logging.info("Creating Excel table.")
    excel_creator = ExcelCreator('data.xlsx')
    headers = ['Title', 'Date', 'Description', 'Picture Filename', 'Count of Search Phrases', 'Contains Money']
    excel_creator.create_headers(headers)
    
    # Find all articles using XPath
    articles = driver.find_elements(By.XPATH, "//article[@class='article']")

    # Loop through each article
    for article in articles:

        # Find elements within the current article using relative XPaths
        title_element = article.find_element(By.XPATH, ".//h2/a")
        image_element = article.find_element(By.XPATH, ".//div[@class='m']//img")
        description_element = article.find_element(By.XPATH, ".//div[@class='info']//div[@class='content']//p[@class='dek']")
        date_element = article.find_element(By.XPATH, ".//div[@class='info']//header[@class='info-header']//div[@class='meta']//span[@class='time']")

        # Extract data from elements
        title = title_element.text
        image_src = image_element.get_attribute("src")
        description = description_element.text
        date = date_element.text

        # Count occurrences of the search phrase in the title and description
        phrase_counter = (
            FoxNewsSearch.phrase_counter(title, phrase_to_search) +
            FoxNewsSearch.phrase_counter(description, phrase_to_search)
        )
        contains_money = FoxNewsSearch.contains_money(title + description)

        # Download the associated image
        FoxNewsSearch.download_image(image_src, f"img_{i}.jpg")

        # Add the extracted data to the Excel sheet
        data = [title, date, description, f"img_{i}.jpg", phrase_counter, contains_money]
        excel_creator.add_row(data)

        logging.debug("Processed article: %s", title)
        

   logging.info("Completed processing all news articles.")
    
    # Save the Excel file
    excel_creator.save_file()
    logging.info("Excel saved succesfully.")
    # Close the page
    page.close()
    logging.info("RPA task completed.")
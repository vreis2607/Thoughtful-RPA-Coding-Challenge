import os
import re
import time
import logging
from datetime import datetime, timedelta
import requests
from robocorp.tasks import task
from robocorp import workitems
from dateutil import parser
from RPA.Excel.Files import Files
from RPA.Browser.Selenium import Selenium
import string

@task
def WebScraper():
    """Main function to run the news scraping process."""

    # Set up logging to track the script's progress and errors
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    # Retrieve variables from the Work Item
    item = workitems.inputs
    search_phrase = item.current.payload.get('search_phrase')  # Term to search for in the news
    category = item.current.payload.get('category')  # Category to filter news by
    months = int(item.current.payload.get('months'))  # Number of months to consider for news articles

    logging.info(f"Search Phrase - {search_phrase}")
    logging.info(f"Category - {category}")
    logging.info(f"Months - {months}")

    class NewsScraper:
        def __init__(self, search_phrase, category, months):
            """Initialize the scraper with search parameters and set up the WebDriver."""
            self.search_phrase = search_phrase
            self.category = category
            self.months = months
            self.browser = Selenium()  # Create a new Selenium instance
            self.news_items = []  # List to store the news items we scrape

        def setup_driver(self):
            """Set up the WebDriver and open the target website."""
            self.browser.open_available_browser("https://www.latimes.com/")  # Open the Los Angeles Times website
            self.browser.maximize_browser_window()  # Make sure the browser window is full size

        @staticmethod
        def clear_output_directory():
            """Clear the downloadedFiles directory to remove old results."""
            downloaded_files_path = 'output/downloadedFiles'
            if os.path.exists(downloaded_files_path):
                for root, dirs, files in os.walk(downloaded_files_path, topdown=False):
                    for name in files:
                        os.remove(os.path.join(root, name))  # Remove files in the downloadedFiles directory
                    for name in dirs:
                        os.rmdir(os.path.join(root, name))  # Remove subdirectories
                os.rmdir(downloaded_files_path)  # Remove the main downloadedFiles directory
            os.makedirs(downloaded_files_path, exist_ok=True)  # Recreate the downloadedFiles directory

        def open_site(self):
            """Open the site and wait until the search button is available."""
            logging.info("Opening site...")
            self.browser.wait_until_element_is_visible('css:[data-element="search-button"]', timeout=10)  # Wait for the search button to appear

        def retry_action(self, func, retries=3, delay=2):
            """Attempt an action multiple times if it fails, with a delay between attempts."""
            for attempt in range(retries):
                try:
                    return func()  # Try to execute the action
                except Exception as e:
                    logging.error(f"Error on attempt {attempt + 1}: {e}")  # Log the error and retry
                    time.sleep(delay)
            raise Exception(f"Action failed after {retries} retries.")  # Raise an exception if it still fails

        def search_news(self):
            """Perform a search for news articles using the search phrase."""
            logging.info("Searching for news...")
            self.browser.click_element('css:[data-element="search-button"]')  # Click the search button
            search_box = self.browser.find_element('name:q')  # Locate the search box
            self.browser.input_text(search_box, self.search_phrase)  # Type the search phrase into the box
            self.browser.submit_form(search_box)  # Submit the search form

        def filter_by_category(self):
            """Filter the news articles by the specified category."""
            logging.info(f"Filtering by category: {self.category}...")
            try:
                self.browser.click_element(f'xpath://span[text()="{self.category}"]')  # Click the category link
            except Exception as e:
                logging.error(f"Category '{self.category}' not found: {e}")  # Log an error if the category is not found

        def search_newest(self):
            """Sort the search results by the newest articles."""
            logging.info("Sorting by newest...")

            def select_newest():
                """Select 'Newest' from the sort dropdown and verify the selection."""
                sort_by = self.browser.find_element('name:s')  # Find the sort dropdown
                self.browser.select_from_list_by_label(sort_by, "Newest")  # Choose 'Newest' from the list
                time.sleep(2)
                # Verify if 'Newest' was selected
                selected_option = self.browser.find_element('css:option[selected="selected"]')  # Find the selected option
                selected_text = selected_option.text  # Get the text of the selected option
                return selected_text

            # Retry selecting 'Newest' until it succeeds or reach max attempts
            max_retries = 5
            for attempt in range(max_retries):
                selected_text = select_newest()  # Try selecting 'Newest'
                if selected_text == "Newest":
                    logging.info("Successfully sorted by newest.")  # Log success if correct option is selected
                    return
                else:
                    logging.warning(f"Attempt {attempt + 1} failed. Selected text: {selected_text}. Retrying...")  # Log if the attempt failed
                    time.sleep(2)

            logging.error("Failed to sort by newest after multiple attempts.")  # Log an error if all attempts fail
            raise Exception("Failed to sort by newest.")

        def parse_relative_date(self, date_str):
            """Parse relative time descriptions like '30 minutes ago'."""
            now = datetime.now()
            match = re.match(r'(\d+)\s*(minutes?|hours?|days?|weeks?)\s+ago', date_str, re.IGNORECASE)
            if match:
                amount, unit = match.groups()
                amount = int(amount)
                if 'minute' in unit:
                    delta = timedelta(minutes=amount)
                elif 'hour' in unit:
                    delta = timedelta(hours=amount)
                elif 'day' in unit:
                    delta = timedelta(days=amount)
                elif 'week' in unit:
                    delta = timedelta(weeks=amount)
                else:
                    return now  # Default to now if unit is unrecognized
                return now - delta
            return parser.parse(date_str)  # Fallback to parsing the date string

        def within_timeframe(self, date_str):
            """Check if a date string is within the specified timeframe."""
            now = datetime.now()

            # Calculate start date based on the number of months
            if self.months == 0:
                # For 0 months, start from the beginning of the current month
                start_date = now.replace(day=1)
            else:
                # Calculate start date by subtracting months from the current month
                start_date = now.replace(day=1) - timedelta(days=self.months * 30)

            # Define the end date as the start of the next month
            end_date = (now.replace(day=1) + timedelta(days=31)).replace(day=1)

            try:
                date = self.parse_relative_date(date_str)
            except ValueError as e:
                logging.error(f"Error parsing date: {e}")
                return False

            # Check if the date falls within the timeframe
            return start_date <= date < end_date

        def get_news(self):
            """Retrieve news articles and add them to the news_items list."""
            logging.info("Retrieving news items...")
            time.sleep(2)
            pages_text = self.browser.find_element('css:[class="search-results-module-page-counts"]').text  # Get page count text
            total_pages = self.extract_total_pages(pages_text)  # Determine the total number of pages
            for page in range(min(total_pages, 9)):  # Loop through pages, up to a maximum of 10
                articles = self.browser.find_elements('css:[class="promo-wrapper"]')  # Get all articles on the page
                found_valid_item = False
                for article in articles:
                    news_item = self.extract_news_item(article)  # Extract details from each article
                    if news_item and self.within_timeframe(news_item["date"]):
                        self.news_items.append(news_item)  # Add valid news item to the list
                        found_valid_item = True

                if not found_valid_item:
                    logging.info(f"No valid items found on page {page + 1}. Exiting loop.")  # Log if no valid items found
                    break

                next_page_button = self.browser.find_element('css:[class="search-results-module-next-page"]')  # Find 'Next' page button
                try:
                    self.browser.click_element(next_page_button)  # Click to go to the next page
                except Exception as e:
                    logging.error(f"Can't go to the next page: {e}")  # Log error if unable to navigate to the next page

        def extract_total_pages(self, pages_text):
            """Extract the total number of pages from the page count text."""
            pattern = r'(?<=of)(.*)'  # Regex to extract the number of pages
            match = re.search(pattern, pages_text)  # Find match for the pattern
            pages = match.group(1).strip()  # Clean the extracted page count
            pages = pages.translate(str.maketrans('', '', string.punctuation))  # Remove punctuation
            return int(pages) - 1  # Convert to integer and adjust for zero-based index

        def extract_news_item(self, article):
            """Extract details from a single news article."""
            try:
                title = article.find_element('css selector', '.promo-title').text  # Get article title
            except:
                title = ""
            try:
                description = article.find_element('css selector', '.promo-description').text  # Get article description
            except:
                description = ""
            try:
                image_url = article.find_element('css selector', '.image').get_attribute("src")  # Get image URL
            except:
                image_url = ""
            try:
                date = article.find_element('css selector', '.promo-timestamp').text  # Get article date
            except:
                date = ""
            return {"title": title, "description": description, "image_url": image_url, "date": date}

        def sanitize_filename(self, filename):
            """Remove invalid characters from filenames and ensure it's within length limits."""
            filename = re.sub(r'[\\/*?:"<>|]', "_", filename)  # Remove invalid characters
            return filename[:255]  # Limit filename length to avoid file system issues

        def download_image(self, url, path):
            """Download an image from the URL and save it to the given path."""
            try:
                response = requests.get(url, stream=True)
                response.raise_for_status()  # Raise an error for HTTP issues
                with open(path, 'wb') as file:
                    for chunk in response.iter_content(1024):
                        file.write(chunk)
            except requests.RequestException as e:
                logging.error(f"Failed to download image from {url}: {e}")  # Log any errors encountered

        def save_news(self):
            """Save the collected news items to an Excel file."""
            logging.info("Saving news items to Excel file...")
            excel = Files()
            excel.create_workbook("output/news_items.xlsx")  # Create a new Excel workbook
            sheet_name = "News_Items"
            headers = ["Title", "Date", "Description", "Image Filename", "Search Count", "Contains Money"]
            
            excel.rename_worksheet("Sheet", sheet_name)  # Rename the default sheet
            excel.append_rows_to_worksheet([headers], sheet_name)  # Add headers to the worksheet

            rows = []
            for item in self.news_items:
                title = item["title"]
                date = item["date"]
                description = item["description"]
                
                image_url = item["image_url"]
                if image_url:
                    image_filename = self.sanitize_filename(image_url.split('/')[-1])  # Clean up the image filename
                else:
                    image_filename = ""
                
                search_count = title.lower().count(self.search_phrase.lower()) + description.lower().count(self.search_phrase.lower())  # Count occurrences of search phrase
                contains_money = bool(re.search(r"\$\d+(?:\.\d{1,2})?|\d+\s?(?:USD|dollars)", title + " " + description))  # Check if money is mentioned

                if image_url:
                    image_path = os.path.join('output/downloadedFiles', image_filename)
                    self.download_image(image_url, image_path)  # Download the image

                row_data = [title, date, description, image_filename, search_count, contains_money]
                rows.append(row_data)

            excel.append_rows_to_worksheet(rows, sheet_name)  # Add news items to the worksheet
            excel.save_workbook()  # Save the Excel file

    # Create and run the NewsScraper instance
    scraper = NewsScraper(search_phrase, category, months)
    scraper.clear_output_directory()
    scraper.setup_driver()
    scraper.open_site()
    scraper.retry_action(scraper.search_news)
    scraper.filter_by_category()
    scraper.search_newest()
    scraper.get_news()
    scraper.save_news()
    logging.info("News scraping completed successfully.")

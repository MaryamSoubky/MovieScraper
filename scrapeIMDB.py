from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
import time
import random
import openpyxl

# Configuration
CHROME_DRIVER_PATH = r'C:\Users\soubk\OneDrive\Desktop\chromedriver-win64\chromedriver.exe'
OUTPUT_FILE = 'IMDB_Top_250_Movies_With_Genres.xlsx'

def setup_driver():
    """Configure Chrome with anti-detection settings"""
    chrome_options = Options()
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
    service = Service(executable_path=CHROME_DRIVER_PATH)
    return webdriver.Chrome(service=service, options=chrome_options)

def initialize_excel():
    """Initialize and configure the Excel workbook"""
    excel = openpyxl.Workbook()
    sheet = excel.active
    sheet.title = 'IMDB Top 250 Movies'
    sheet.append(['Rank', 'Name', 'Year', 'Rating', 'Genre'])
    return excel

def scroll_page(driver):
    """Scroll to the bottom of the page to load all content"""
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(random.uniform(1.5, 3))
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height

def get_movie_links(soup):
    """Extract and return movie links from the main page"""
    movie_links = []
    movies_list = soup.find('ul', class_="ipc-metadata-list ipc-metadata-list--dividers-between sc-e22973a9-0 khSCXM compact-list-view ipc-metadata-list--base")
    
    if movies_list:
        for movie in movies_list.find_all('li', class_="ipc-metadata-list-summary-item"):
            link = movie.find('a', class_="ipc-title-link-wrapper")
            if link:
                full_link = "https://www.imdb.com" + link['href'].split('?')[0]
                movie_links.append(full_link)
    return movie_links

def extract_movie_info(movie_item, index):
    """Extract basic movie info from the list item"""
    title_element = movie_item.find('a', class_="ipc-title-link-wrapper")
    name = title_element.text.strip().split('.', 1)[-1].strip()
    rank = title_element.text.strip().split('.')[0].strip()
    
    metadata = movie_item.find('div', class_="cli-title-metadata")
    year = metadata.find('span').text.strip() if metadata else "N/A"
    
    rating = movie_item.find('span', class_="ipc-rating-star--rating")
    rating_value = rating.text.strip() if rating else "N/A"
    
    return {
        'rank': rank,
        'name': name,
        'year': year,
        'rating': rating_value,
        'index': index
    }

def scrape_genres(driver, url):
    """Robust genre scraping with multiple fallback methods"""
    try:
        driver.get(url)
        
        # Wait for either the modern or legacy genre element
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "[data-testid='genres'], .ipc-chip-list--base"))
            )
        except TimeoutException:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(),'Genre') or contains(text(),'genres')]"))
            )

        # Multiple selectors to find genres
        selectors = [
            "[data-testid='genres'] a",  # Modern IMDb
            ".ipc-chip-list--base a",    # Newer variant
            ".see-more.inline a[href*='genre']",  # Legacy
            "a[href*='genres=']"         # Fallback
        ]
        
        for selector in selectors:
            elements = driver.find_elements(By.CSS_SELECTOR, selector)
            if elements:
                return ', '.join([el.text for el in elements if el.text])
                
        return "N/A"
        
    except Exception as e:
        print(f"Error scraping {url}: {str(e)[:200]}")
        return "N/A"
    finally:
        time.sleep(random.uniform(2, 5))

def main():
    excel = initialize_excel()
    sheet = excel.active
    driver = setup_driver()
    
    try:
        # Open IMDb Top 250
        print("Loading IMDb Top 250 page...")
        driver.get('https://www.imdb.com/chart/top/')
        
        # Wait for page to load
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CLASS_NAME, "ipc-metadata-list"))
        )
        
        # Scroll to load all movies
        print("Scrolling to load all movies...")
        scroll_page(driver)
        
        # Parse page
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        movie_links = get_movie_links(soup)
        movies_list = soup.find_all('li', class_="ipc-metadata-list-summary-item")
        
        print(f"Found {len(movie_links)} movies. Starting genre scraping...")
        
        for i, link in enumerate(movie_links[:250]):  # Process all 250 movies
            try:
                movie_info = extract_movie_info(movies_list[i], i)
                genres = scrape_genres(driver, link)
                
                sheet.append([
                    movie_info['rank'],
                    movie_info['name'],
                    movie_info['year'],
                    movie_info['rating'],
                    genres
                ])
                
                print(f"Processed {i+1}/250: {movie_info['rank']}. {movie_info['name']} | Genres: {genres}")
                
            except Exception as e:
                print(f"Failed on movie {i+1}: {str(e)[:200]}")
                continue
                
    except Exception as e:
        print(f"Fatal error: {str(e)}")
    finally:
        driver.quit()
        excel.save(OUTPUT_FILE)
        excel.close()
        print(f"Scraping complete. Data saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
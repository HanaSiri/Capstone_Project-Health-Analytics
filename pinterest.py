from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
import urllib.parse

def scrape_pinterest_with_selenium(hashtag):
    # Properly encode the hashtag for the URL 
    encoded_hashtag = urllib.parse.quote(f'#{hashtag}')

    # Use webdriver_manager to manage Chromedriver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.get(f"https://www.pinterest.com/search/pins/?q={encoded_hashtag}")

    # Wait for content to load
    try:
        WebDriverWait(driver, 100).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, '[data-test-id="pin"]'))
        )
    except Exception as e:
        print(f"Error waiting for elements for hashtag {hashtag}: {e}")
        driver.quit()
        return []

    # Scroll down to load more pins 
    SCROLL_PAUSE_TIME = 10
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(SCROLL_PAUSE_TIME)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height

    pins = []
    try:
        elements = driver.find_elements(By.CSS_SELECTOR, '[data-test-id="pin"]')
        print(f"Found {len(elements)} elements for hashtag: {hashtag}")

        for element in elements:
            try:
                img_tag = element.find_element(By.TAG_NAME, 'img')
                title = img_tag.get_attribute('alt')
                image_url = img_tag.get_attribute('src')
                img_description = img_tag.get_attribute('text')
                print(f"Title: {title}, Image URL: {image_url}, Image Description: {img_description}")
                pins.append({'Hashtag': hashtag,'Title': title, 'Image URL': image_url, 'Image Description': img_description})
            except Exception as e:
                print(f"Error extracting data from element: {e}")
                continue
    except Exception as e:
        print(f"Error finding elements for hashtag {hashtag}: {e}")
    
    driver.quit()
    return pins

# Hashtags to scrape
hashtag = ['CoffeeFruit', 'PlantBased', 'ProteinPowder', 'HealthSupplements','AntiAging','BrainHealth','Superfoods','HealthyLiving','NaturalRemedies','Nutrition',
           'Wellness','AntiAging','HealthAndWellness','VeganSupplements','BrainBoost','OrganicLiving','HolisticHealth','CleanEating','SuperfoodPowder','MindBodySoul',
           'YouthfulSkin','FitnessNutrition','Longevity','WholeFoodSupplements','CognitiveHealth','HealthyAging']

# Collect data
all_data = []
for tag in hashtag:
    print(f"Scraping data for hashtag: {tag}")
    data = scrape_pinterest_with_selenium(tag)
    all_data.extend(data)
    
# Create DataFrame
df = pd.DataFrame(all_data)

# Save to XLSX
df.to_excel('pinterest_data_selenium_3.xlsx')

print("Data scraping complete and saved to 'pinterest_data_selenium_3.xlsx'.")
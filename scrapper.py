import os
import urllib.parse
import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ========== Configuration ==========
search = urllib.parse.quote(input("Enter the search (e.g., hospitals in Kasaragod): "))
chromedriver_path = "/usr/bin/chromedriver"  # Replace with your actual path
os.environ["PATH"] += os.pathsep + os.path.dirname(chromedriver_path)

# Optional: Run headless if you don’t want browser popping up
options = Options()
# options.add_argument("--headless")

# ========== Start Driver ==========
driver = webdriver.Chrome(options=options)

# ========== Step 1: Go to Google Maps Search Page ==========
url = f"https://www.google.com/maps/search/{search}"
driver.get(url)
input("Press Enter after map results have loaded...")

# ========== Step 2: Get All Search Result Links ==========
soup = BeautifulSoup(driver.page_source, 'html.parser')
main_links = [a['href'] for a in soup.find_all("a", {"class": "hfpxzc"}) if a.get('href')]

print(f"Found {len(main_links)} links")

# ========== Step 3: Visit Each Result Page ==========
final_data = []

for index, link in enumerate(main_links):
    driver.get(link)
    time.sleep(5)  # Allow full content to load

    # Try to get hospital name
    try:
        name_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "h1.DUwDvf.fontHeadlineLarge"))
        )
        name = name_element.text
    except Exception:
        try:
            # Fallback with different tag
            name = driver.find_element(By.TAG_NAME, "h1").text
        except Exception:
            name = "Name Not Found"

    # Try to get phone number
    phone = ""
    try:
        buttons = driver.find_elements(By.CLASS_NAME, "CsEnBe")
        for b in buttons:
            if b.get_attribute("aria-label") and "Phone" in b.get_attribute("aria-label"):
                phone = b.get_attribute("aria-label").split(":")[1].strip()
                break
    except:
        phone = ""

    # Get other details using BeautifulSoup
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    details_raw = soup.find_all("div", {"class": "rogA2c"})
    details = ", ".join([d.text for d in details_raw])

    # Append to list
    final_data.append({
        "name": name,
        "phone": phone,
        "details": details
    })

    print(f"[{index+1}/{len(main_links)}] Saved: {name}")

# ========== Step 4: Save to Excel ==========
df = pd.DataFrame(final_data)
df.to_excel("data.xlsx", index=False)
print("Data saved to data.xlsx")

driver.quit()


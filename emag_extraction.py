from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
import time, random, pandas as pd, os

url = "https://www.emag.ro/vendors/vendor/xtremelighting/p{}"
from_page = 1
to_page = 29

results = []
collected = []

def get_product_data(product_url, page):
    data = {}

    driver.execute_script("window.open('{}')".format(product_url))
    time.sleep(random.uniform(1,2))
    driver.switch_to.window(driver.window_handles[-1])

    wait = WebDriverWait(driver, 20)
    wait.until(EC.visibility_of_element_located((By.CLASS_NAME, 'highlight-box')))

    data["product_page"] = page
    data["product_url"] = product_url

    product_name = driver.find_element(By.CLASS_NAME, "page-title").text
    data["product_name"] = product_name

    product_code = driver.find_element(By.CLASS_NAME, "product-code-display").text.replace("Cod produs: ", "")
    data["product_code"] = product_code

    # time.sleep(random.uniform(1,2))

    product_price = driver.find_element(By.CLASS_NAME, "highlight-box").find_element(By.CLASS_NAME, "product-new-price").text
    data["product_price"] = product_price

    # time.sleep(random.uniform(1,2))

    product_images = [
        x.get_attribute("href") for x in driver.find_element(By.CLASS_NAME, "ph-scroller").find_elements(By.CLASS_NAME, "product-gallery-image")
    ]
    data["product_images"] = ",".join(product_images)

    # time.sleep(random.uniform(1,2))
    
    try:
        desription_body = driver.find_element(By.CLASS_NAME, "product-page-description-text")
        driver.execute_script("arguments[0].scrollIntoView();", desription_body)
        try:
            see_more = desription_body.find_element(By.XPATH, "//a[@data-phino='CollapseOne' and @data-ph-target='#description-body']")
            see_more.click()
            time.sleep(random.uniform(1,2))
        except Exception:
            print("See more description - Element not found for url {}".format(product_url))
        finally:
            product_description = desription_body.text
            data["product_description"] = product_description
    except Exception:
        data["product_description"] = ""

    # time.sleep(random.uniform(1,2))

    try:
        specifications_body = driver.find_element(By.ID, "specifications-body")
        driver.execute_script("arguments[0].scrollIntoView();", specifications_body)
        try:
            see_more = specifications_body.find_element(By.XPATH, "//button[@data-phino='CollapseOne' and @data-ph-target='#specifications-body']")
            see_more.click()
            time.sleep(random.uniform(1,2))
        except Exception:
            print("See more specifications - Element not found for url {}".format(product_url))
        finally:
            specs = specifications_body.find_element(By.TAG_NAME, "tbody").find_elements(By.TAG_NAME, "tr")
            for spec in specs:
                row = spec.find_elements(By.TAG_NAME, "td")
                spec_title = row[0].text
                spec_value = row[1].text
                data[spec_title] = spec_value
    except Exception:
        pass

    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    collected.append(product_url)
    return data

driver = webdriver.Chrome("chromedriver.exe")

for page in range(from_page, to_page + 1):
    page_url = url.format(page)
    driver.get(page_url)

    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, 'card_grid')))

    time.sleep(random.uniform(1,3))

    products_grid = driver.find_element(By.ID, "card_grid")

    products = products_grid.find_elements(By.XPATH, "./div[contains(@class, 'card-item') and contains(@class, 'card-standard') and contains(@class, 'js-product-data')]")

    links = [p.find_element(By.TAG_NAME, "a").get_attribute("href") for p in products]

    for link in links:
        if link not in collected:
            time.sleep(random.uniform(1,3))
            res = get_product_data(link, page)
            if bool(res):
                results.append(res)

all_keys = set().union(*results)
for d in results:
    missing_keys = all_keys - d.keys()
    for key in missing_keys:
        d[key] = ''

df = pd.DataFrame(results)
df.to_excel("emag_extracted_pages_remaining.xlsx")
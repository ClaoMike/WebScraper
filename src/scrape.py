import os
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import csv
from selenium.webdriver.chrome.options import Options

# Settings
PRINT_TO_CONSOLE = True
HIDE_BROWSER = False

# input and output files
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(BASE_DIR)
INPUT_DIR = os.path.join(PROJECT_ROOT, 'data', 'input')
OUTPUT_DIR = os.path.join(PROJECT_ROOT, 'data', 'output')

input_file_name = 'DeviceModel.txt'
output_file_name = 'DeviceDetails.csv'

input_file_path = os.path.join(INPUT_DIR, input_file_name)
output_file_path = os.path.join(OUTPUT_DIR, output_file_name)

# the search page URL
BASE_URL = "https://psref.lenovo.com/search/"

# chrome comes with selenium
chrome_options = Options()
if HIDE_BROWSER:
    chrome_options.add_argument("--headless")  # Run in headless mode

driver = webdriver.Chrome(options=chrome_options)
driver.get(BASE_URL)

def search_ID(id):
    if PRINT_TO_CONSOLE:
        print(f"Searching for: {id}")

    search_box = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "as_quick_search_text"))
    )
    search_box.clear()
    search_box.send_keys(id)
    search_box.send_keys(Keys.RETURN)

def print_device_details(device_ID, device_name, device_link, device_processor, device_memory, device_storage):
    if PRINT_TO_CONSOLE:
        print("\n=====================================================================================")
        print(f"Device ID: {device_ID}")
        print(f"Device Name: {device_name}")
        print(f"Product link: {device_link}")
        print(f"Device Processor: {device_processor}")
        print(f"Device Memory: {device_memory}")
        print(f"Product Storage: {device_storage}")
        print("=====================================================================================\n")

with open(output_file_path, 'w', newline='', encoding='utf-8') as csvfile:
    csv_writer = csv.writer(csvfile)
    csv_writer.writerow(["ID", "Name", "Link", "Processor", "Memory", "Storage"])

try:
    with open(input_file_path, 'r') as file, open(output_file_path, 'a', newline='', encoding='utf-8') as csvfile:
        # the writer
        csv_writer = csv.writer(csvfile)

        # for each line, remove leading/trailing whitespace and newline characters and skip empty lines
        for line in file:
            device_ID = line.strip()
            if not device_ID:
                continue

            # default values
            device_name = "N/A"
            device_link = "N/A"
            device_processor = "N/A"
            device_memory = "N/A"
            device_storage = "N/A"

            # search for the ID
            search_ID(device_ID)

            # wait for the results
            time.sleep(1)

            # Wait for the product_array div
            product_array = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "product_array"))
            )

            # Locate all item_product_body elements within the product_array
            item_bodies = product_array.find_elements(By.CLASS_NAME, "item_product")

            for item_body in item_bodies:
                # locate the title
                a_element = item_body.find_element(
                    By.CLASS_NAME, "productContent_list_modets"
                ).find_element(
                    By.TAG_NAME, "h2"
                ).find_element(
                    By.TAG_NAME, "a"
                )

                # extract the text and href from the <a> element
                device_name = a_element.text  # extract visible text
                device_link = a_element.get_attribute("href")  # extract href

                # Locate all <p> tags inside modets_list_cur
                paragraphs = item_body.find_element(
                    By.CLASS_NAME, "modets_B"
                ).find_element(
                    By.CLASS_NAME, "modets_list_cur"
                ).find_elements(
                    By.TAG_NAME, "p"
                )

                # extract processor, memory and storage
                for paragraph in paragraphs:
                    try:
                        b_element = paragraph.find_element(By.TAG_NAME, "b")  # locate the <b> element
                        key = b_element.text.strip()  # get the <b> text
                        value = paragraph.text.replace(b_element.text, "").strip()  # extract text after <b>

                        if key == "Processor:":
                            device_processor = value
                        elif key == "Memory:":
                            device_memory = value
                        elif key == "Storage:":
                            device_storage = value

                    except:
                        continue  # skip paragraphs without <b> elements

                print_device_details(device_ID, device_name, device_link, device_processor, device_memory, device_storage)

            csv_writer.writerow([device_ID, device_name, device_link, device_processor, device_memory, device_storage])

except FileNotFoundError:
    print(f"File {input_file_name} not found in {INPUT_DIR}.")
finally:
    driver.quit()
"""BEFORE RUNNING THE SCRIPT, FIRST READ THE DOCUMENTATION FILE"""

# Importing libraries

import os
import re
import json
import docx
import pandas as pd

# Importing Selenium 

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# Function to setup selenium driver and chrome driver location
def setup_selenium_driver():
    options = Options()
    options.headless = True
    service = Service(executable_path = r"C:\Users\ayush\Downloads\chromedriver-win64\chromedriver.exe")
    driver = webdriver.Chrome(service = service, options = options)
    return driver

# Function to extract urls from the .docx file
def extract_content_from_docx(file_path):
    if not os.path.exists(file_path):
        print("File does not exist.")
        return []
    
    doc = docx.Document(file_path)
    content = []
    # Url pattern recognition
    url_regex = r'https?://(?:[-\w.]|(?:%[\da-fA-F]{2}))+'
    for para in doc.paragraphs:
        urls = re.findall(url_regex, para.text, re.IGNORECASE)
        if urls:
            content.append({'urls': urls})
    return content

# Function to extract data from a given URL using Selenium
def extraction(url, driver, count):
    try:
        # get request to get the url
        driver.get(url)

        print(f'Data fetching Started from url No : {count} and url : {url}')
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'body')))
        
        # text extraction
        text = " ".join([element.text for element in driver.find_elements(By.CSS_SELECTOR, 'p, h1, h2, h3, h4, h5, h6')])
        text = text.replace('\n', '').replace('\t', '')
        text = text.strip()
        # Replace or remove illegal characters for Excel
        text = re.sub(r'[\x00-\x1F\x7F-\x9F]', '', text) # Regex to remove control characters

        # image link extraction
        image_urls = [img.get_attribute('src') for img in driver.find_elements(By.TAG_NAME, 'img')]
        
        # web link extraction
        web_links = [link.get_attribute('href') for link in driver.find_elements(By.TAG_NAME, 'a') if link.get_attribute('href') and link.get_attribute('href').startswith('https')]
        
        print(f'Data fetched successfully from url: {url}')
        print('*'*125)
        
        return {
            'URL': url, 
            'TEXT': text, 
            'IMAGE LINKS': ', '.join(image_urls), 
            'WEB LINKS':  ', '.join(web_links)
        }
    except Exception as e:
        print(f"An error occurred while extracting content from {url} : {e}")
        print('*'*125)
        return None

# Function to process document content
def process_document_content(content, driver):
    data = []
    for item in content:
        for url in item['urls']:
            info = extraction(url, driver, len(data) + 1)
            if info:
                data.append(info)
    return data

# Function to save the data in json and excel file format
def save_data(data, output_dir):
    json_path = os.path.join(output_dir, 'output.json')
    excel_path = os.path.join(output_dir, 'output.xlsx')

    with open(json_path, 'w') as json_file:
        json.dump(data, json_file)
    print(f"Data saved to {json_path}")

    if data:
        df = pd.DataFrame(data)
        df.to_excel(excel_path, index=False)
        print(f"Data saved to {excel_path}")

# Main Function 
def main():
    # Setup driver
    driver = setup_selenium_driver()
    
    # Process document and path of the document location
    docx_file_path = r"C:\Users\ayush\Downloads\Fresher\Python2_Python\Python_Assigment_2.docx"
    document_content = extract_content_from_docx(docx_file_path)
    data = process_document_content(document_content, driver)
    
    # Cleanup / Exiting from the chromedriver
    driver.quit()
    
    # Directory Setup
    output_dir = r'C:\Users\ayush\Downloads\Fresher\Python2_Python'
    os.makedirs(output_dir, exist_ok=True)
    save_data(data, output_dir)

    print("Scraping complete and data has been saved to json and excel file.")

if __name__ == "__main__":
    main()
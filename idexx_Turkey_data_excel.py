#new file
import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.service import Service
from selenium.common.exceptions import ElementClickInterceptedException
from webdriver_manager.firefox import GeckoDriverManager
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
import time
import requests
import PyPDF2
import io



selected_folder_path = None  # Initialize the variable to store the selected folder path


    

def choose_and_return_file():
    global selected_folder_path
    selected_folder_path = filedialog.askopenfilename()
    root.destroy()  # Close the Tkinter window after selecting a file
    print(file_path)
    return file_path


def check_words_in_pdf(url):
    print("in check words in pdf function")
    response = requests.get(url)
    all_words = []
    words_to_check = ['P04-006','TS EN 13043','TS 706 EN', 'TS EN ISO 9712','TS EN ISO 7899','TS EN ISO 7899-2','TS EN ISO 9308','TS EN ISO 6892','ISO 7899','ISO 9308','SMWW 9223','SMWW 9221','SMWW 9222','SMWW 9221','ISO 7899','SMWW 9230','ISO 1626','SMWW 9230','ISO 6222','SMWW 9215','ISO 1626','ISO 1173','ASTMD 842921']
    if response.status_code == 200:
        print("in search process")
        # Parse the PDF content using PyPDF2
        pdf_file = io.BytesIO(response.content)
        pdf_reader = PyPDF2.PdfFileReader(pdf_file)

        # Initialize a dictionary to store the words found and their cases
        words_found = {}

        # Search for the desired words in the PDF
        for word in words_to_check:
            for page_num in range(pdf_reader.getNumPages()):
                page = pdf_reader.getPage(page_num)
                text = page.extractText()
                if word.lower() in text.lower():
                    if word.islower() and word.lower() in text:
                        words_found[word] = "lowercase"
                    elif word.isupper() and word.upper() in text:
                        words_found[word] = "uppercase"
        if words_found:
            print("words found fetching..")
            for word, case in words_found.items():
                all_words.append(word)
                
        return all_words
    else:
        print("Failed to fetch the PDF content.")
        words_found = "not found"
        return words_found
    
    
def pdf_link(link):
    print("URL: ",link)
    try:
        driver = webdriver.Chrome(ChromeDriverManager().install())
        driver.get(link)
        
    except:
        try:
            time.sleep(2)
            driver = webdriver.Chrome(ChromeDriverManager().install())
            driver.get(link)
            time.sleep(2)
        except:
            print("except")
    try:
        time.sleep(2)
        pdf_button = driver.find_element(By.XPATH, '/html/body/app-root/div[1]/esign-validator/div/div/a[1]')
        anchor_tags = driver.find_elements(By.TAG_NAME, 'a')
        print("in anchor_tags")
        for tag in anchor_tags:
            print("in for loop anchor_tags:")
            href = tag.get_attribute('href')  # Get the URL of the anchor tag
            text = tag.text  # Get the visible text of the anchor tag
            print(f"Anchor tag: href={href}, text={text}")
            driver.get(href)
            time.sleep(1)
            break
        print("out of loop")
        all_words = check_words_in_pdf(href)
        driver.quit()
        print(all_words)
        return all_words
    
    except:
        print("exception at finding url for pdf")
        all_words = ["not found"]
        return all_words



def find_matching_words(url):
    # Send a GET request to the website
    #target_words = ['TS EN 13043','TS 706 EN', 'Blok Kalibratör', 'TS EN ISO 9712','TS EN ISO 7899','TS EN ISO 7899-2','TS EN ISO 9308','TS EN ISO 6892','ISO 7899','ISO 9308','SMWW 9223','SMWW 9221','SMWW 9222','SMWW 9221','ISO 7899','SMWW 9230','ISO 1626','SMWW 9230','ISO 6222','SMWW 9215','ISO 1626','ISO 1173','ASTMD 842921']
    target_words = ['9308','9308-1','9308-2','9308-3','9308-4','9308-5','9308-6','9308-7','9308-8','9308-9','9308-10','9308-11','9308-12','9230','9230 B','9230 D','6222','9215','16266','16266-2','11731','8429',' 8429-21']
    print("working...")
    response = requests.get(url)

    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(response.text, 'html.parser')

    # Find all text occurrences on the website
    all_text = soup.get_text()
    
    # Convert the target_words list to lowercase for case-insensitive matching
    target_words = [word.lower() for word in target_words]

    # Check if any of the target words are present in the website text
    matching_words = []
    for word in target_words:
        if word in all_text.lower():
            matching_words.append(word)
            print("found searching...")

    return matching_words


# Example retrieve_data function (replace with your actual function)
def retrieve_data(text):
    name_pattern = re.compile(r'\n.*\n(.*?)\n')
    name_match = re.search(name_pattern, text)
    if name_match:
        name = name_match.group(1)
    else:
        name = None

    # Find the email using regex
    email_pattern = re.compile(r'E-posta: (.*?)\n')
    email_match = re.search(email_pattern, text)
    if email_match:
        email = email_match.group(1)
    else:
        email = None

    # Find the phone number using regex
    phone_pattern = re.compile(r'Telefon: (.*?) Fax:', re.DOTALL)
    phone_match = re.search(phone_pattern, text)
    if phone_match:
        phone_number = phone_match.group(1)
    else:
        phone_number = None

    print("Name:", name)
    print("Email:", email)
    print("Phone Number:", phone_number)
    return name, email, phone_number



#_____________________________________________________________________________


# Create the main window
print("start")
for i in range(3):
    try:
        print("loop try")
        root = tk.Tk()
        root.title("File Chooser")

        # Set the window size
        root.geometry("400x200")  # Width x Height

        # Create a button to choose a file
        choose_button = tk.Button(root, text="Choose previous excel File", command=choose_and_return_file)
        choose_button.pack(pady=20)

        # Start the Tkinter event loop
        root.mainloop()
        print("root main loop")
        print(selected_folder_path)
        df = pd.read_excel(selected_folder_path)
        break
    except:
        print("exception")
        
        
    

#_____________________________________________________________________________

for index, row in df.iterrows():
    if row['Company details']:  # Assuming 'Company Details' is not empty
        name, email, phone_number = retrieve_data(row['Company details'])
        df.at[index, 'name'] = name
        df.at[index, 'email'] = email
        df.at[index, 'phone_number'] = phone_number
df_cleaned = df.dropna()
# print(df_cleaned)

for index, row in df_cleaned.iterrows():
    url = row['Urls']
    matching_words = find_matching_words(url)
    df_cleaned.at[index, 'matching_words'] = ', '.join(matching_words)

print(df_cleaned)


for index, row in df_cleaned.iterrows():
    url = row['Urls']
    if url.endswith("AccreditationCertificate"):
        matching_words = pdf_link(url)
        df_cleaned.at[index, 'matching_words'] = ', '.join(matching_words)

print(df_cleaned)



# df = df_cleaned

# Define the list of predefined values
predefined_values = ['9308','9308-1','9308-2','9308-3','9308-4','9308-5','9308-6','9308-7','9308-8','9308-9','9308-10','9308-11','9308-12','9230','9230 B','9230 D','6222','9215','16266','16266-2','11731','8429',' 8429-21']

# Create new columns based on predefined values and mark 'x' if found
for value in predefined_values:
    df_cleaned[value] = df_cleaned['matching_words'].apply(lambda x: 'x' if value in x else '')

# Display the resulting DataFrame
print(df_cleaned)

root = tk.Tk()
root.title("Folder Chooser")

# Set the window size
root.geometry("400x200")  # Width x Height

# Create a button to choose a folder
choose_button = tk.Button(root, text="Choose Folder", command=choose_and_return_folder)
choose_button.pack(pady=20)

# Start the Tkinter event loop
root.mainloop()

# Print the selected folder path (after the Tkinter event loop has finished)
if selected_folder_path:
    print("Selected Folder:", selected_folder_path)
else:
    print("No folder selected")

file_path = selected_folder_path + '/tukey_df_cmd.xlsx'


df_cleaned.to_excel(file_path, index=True)
print("excel saved")
import time
import pandas as pd
import urllib.parse 
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager 
import mysql.connector
from mysql.connector import Error
from datetime import datetime
from pynput.keyboard import Controller
import ssl

ssl._create_default_https_context = ssl._create_unverified_context


def addCampaign(name):
    query = """ 
    INSERT INTO campaigns (title)
    VALUES (%s)
    """ 
    cursor.execute(query, (name,))
    conn.commit()
    print(f"Campaign '{name}' created.")


def printCampaigns():
    query = """
    SELECT id, title FROM campaigns
    """
    cursor.execute(query)
    results = cursor.fetchall()

    campaigns_dict = {row[0]: row[1] for row in results}
    return campaigns_dict
    
def logMessages(mobile, campaign_id):
    now = datetime.now()
    query = """
    INSERT INTO messages (time_sent, date_sent, mobile, campaign_id)
    VALUES (%s, %s, %s, %s)
    """
    values = (now.strftime('%H:%M:%S'), now.strftime('%Y-%m-%d'), mobile, campaign_id)
    cursor.execute(query, values)
    conn.commit()
    print(f"Log added for {name} ({mobile}) in messages table with campaign ID {campaign_id}.")


####################################################################

try:
    conn = mysql.connector.connect(
        host='127.0.0.1',
        user='root',
        password='Password123',
        database='whatsApp_system'
    )
    cursor = conn.cursor()
except Error as e:
    print(f"Error while connecting to database: {e}")
    exit()
    
print('\n\nWelcome! Choose an option from the below')

selected_campaign_id = None
selected_campaign_title = None

while True:
    print('\n(1) To enter a new Campaign.\n(2) To choose an existing Campaign.\n(0) Exit.')
    ans = input("Enter your choice (1, 2, 0): ")

    if ans == "0":
        print("Exiting program.")
        exit(0)

    elif ans == "1":
        while True:
            newCamp = input("Enter the name of the new Campaign (0 to exit): ")
            if newCamp == "0":
                break
            newCampConf = input("Confirm the name of your new Campaign: ")
            if newCamp == newCampConf:
                addCampaign(newCamp)
                break
            else:
                print("Campaign names do not match. Please try again.\n")

    elif ans == "2":
        campaigns = printCampaigns()

        if not campaigns:
            print("\n----------------")
            print("No Campaigns in database!")
            print("----------------")
            continue

        print("\nExisting Campaigns:\n----------------")
        for cid, title in campaigns.items():
            print(f"ID: {cid}, Title: {title}")
        print("----------------")

        selected_id = input("Select a Campaign by ID(0 to go back): ")

        if selected_id.isdigit():
            selected_id = int(selected_id)
            if selected_id in campaigns:
                selected_campaign_id = selected_id
                selected_campaign_title = campaigns[selected_id]
                print(f"\nYou selected Campaign ID {selected_campaign_id}: {selected_campaign_title}")
                break
            else:
                print("Invalid Campaign ID selected. Returning to main menu.")
        else:
            print("Invalid input. Returning to main menu.")
    else:
        print("Invalid option. Please enter 0, 1, or 2.\n")


###############################################################
file_path = "C:/Users/VB/Desktop/Whatsapp-Project/test.xlsm"
###############################################################


df = pd.read_excel(file_path, engine="openpyxl")
print("Columns in Excel file:", df.columns.tolist())

df.columns = ["Phone#", "IV", "Person in Charge", "Total OS", "Subject"]
df = df[["Phone#", "IV", "Person in Charge", "Total OS"]]

df["Phone#"] = df["Phone#"].astype(str).str.strip()


options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)

MAX_RETRIES = 3
for attempt in range(MAX_RETRIES):
    try:
        service = Service(ChromeDriverManager().install())
        break
    except Exception as e:
        print(f"Attempt {attempt+1} failed: {e}")
        if attempt < MAX_RETRIES - 1:
            time.sleep(5)
        else:
            raise
        
driver = webdriver.Chrome(service=service, options=options)

WHATSAPP_WEB_URL = "https://web.whatsapp.com/"
driver.get(WHATSAPP_WEB_URL)
input("Press Enter after scanning QR Code in WhatsApp Web...")
time.sleep(3)

keyboard = Controller()
keyboard.press("\n")
keyboard.release("\n")
time.sleep(2)

for index, row in df.iterrows():
    phone = row["Phone#"]
    name = row["Person in Charge"]
    prename = row["IV"]
    amount = row["Total OS"]
###############################################################
    message = """\
    Dear {name},
    This is a reminder for the {prename} {amount} OS.
 """
###############################################################

    encoded_message = urllib.parse.quote(message)
 
    whatsapp_url = f"https://web.whatsapp.com/send?phone={phone}&text={encoded_message}"
    driver.get(whatsapp_url)
    time.sleep(10)  

    try:
        send_button = driver.find_element(By.XPATH, '//button[@aria-label="Send"]') 
        send_button.click()
        print(f"Message sent to {name} ({phone})")
        logMessages(phone, selected_campaign_id)
    except Exception as e:
        print(f"Failed to send message to {name} ({phone}). Error: {e}")

    time.sleep(5) 

# Close the browser
driver.quit()

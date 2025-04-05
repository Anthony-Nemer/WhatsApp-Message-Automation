import time
import pandas as pd
import urllib.parse  # For encoding messages in URLs
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import mysql.connector
from mysql.connector import Error
from datetime import datetime
from pynput.keyboard import Controller
 

# pip install pandas openpyxl selenium webdriver-manager

def logMessages(toMobileNumber, amount, name, prename):
    now = datetime.now()
    query = """
    INSERT INTO messages (time_sent, date_sent, toMobileNumber, amount, full_name)
    VALUES (%s, %s, %s, %s, %s)
    """
    values = (now.strftime('%H:%M:%S'), now.strftime('%Y-%m-%d'), toMobileNumber, amount, f"{prename} {name}")
    cursor.execute(query, values)
    conn.commit()
    print(f"Log added for {name} ({toMobileNumber}) in messages table with campaign ID.")


try:
    conn = mysql.connector.connect(
        host='127.0.0.1',
        user='root',
        password='nemeranthony',
        database='whatsApp_system'
    )
    cursor = conn.cursor()
    print("Connected to database successfully.")
except Error as e:
    print(f"Error while connecting to database: {e}")
    exit()

# Load the Excel file
file_path = "C:/Users/VB/Desktop/Whatsapp Project/test.xlsm"
df = pd.read_excel(file_path, engine="openpyxl")
print("Columns in Excel file:", df.columns.tolist())

# Rename and keep only relevant columns
df.columns = ["Phone#", "IV", "Person in Charge", "Total OS", "Subject"]
df = df[["Phone#", "IV", "Person in Charge", "Total OS"]]

# Convert phone numbers to strings (ensure correct format)
df["Phone#"] = df["Phone#"].astype(str).str.strip()

# Initialize Selenium WebDriver
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)  # Keep browser open
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Open WhatsApp Web
WHATSAPP_WEB_URL = "https://web.whatsapp.com/"
driver.get(WHATSAPP_WEB_URL)
input("Press Enter after scanning QR Code in WhatsApp Web...")
keyboard = Controller()
time.sleep(7)  
keyboard.press("\n")  
keyboard.release("\n")
keyboard.press("\n")  
keyboard.release("\n")



# Iterate through the rows and send messages
for index, row in df.iterrows():
    phone = row["Phone#"]
    name = row["Person in Charge"]
    prename = row["IV"]
    amount = row["Total OS"]

    # Format message & encode it properly
    message = """\
        Enter your message here.
        You can include variables like this:    
        Name: {name}
        Amount: {amount}
        """

    encoded_message = urllib.parse.quote(message)
 
    # Open WhatsApp chat
    whatsapp_url = f"https://web.whatsapp.com/send?phone={phone}&text={encoded_message}"
    driver.get(whatsapp_url)
    time.sleep(10)  # Wait for the chat to load

    # Press ENTER to send
    try:
        send_button = driver.find_element(By.XPATH, '//button[@aria-label="Send"]')  # More reliable XPath
        send_button.click()
        print(f"Message sent to {name} ({phone})")
        # logMessages(phone, amount, name, prename)
    except Exception as e:
        print(f"Failed to send message to {name} ({phone}). Error: {e}")

    time.sleep(5)  # Pause before sending the next message

# Close the browser
driver.quit()

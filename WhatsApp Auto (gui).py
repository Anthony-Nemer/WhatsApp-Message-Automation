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
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk 
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# pip install pandas openpyxl selenium webdriver-manager mysql-connector-python

# Database Connection
try:
    conn = mysql.connector.connect(
        host='127.0.0.1',
        user='root',
        password='Password123',
        database='WhatsApp_System'
    )
    cursor = conn.cursor()
    print("Connected to database successfully.")
except Error as e:
    print(f"Error while connecting to database: {e}")
    exit()

def load_users():
    cursor.execute("SELECT id, CONCAT(first_name, ' ', last_name) AS full_name FROM users")
    users = cursor.fetchall()
    return {full_name: user_id for user_id, full_name in users}

def load_campaigns():
    cursor.execute("SELECT id, title FROM campaigns")
    campaigns = cursor.fetchall()
    return {title: campaign_id for campaign_id, title in campaigns}

def add_campaign():
    def save_campaign():
        title = entry_title.get()
        description = entry_description.get()
        user_name = user_dropdown.get()
        
        if not title.strip() or not user_name:
            messagebox.showerror("Error", "All fields are required!")
            return
        
        user_id = users_dict.get(user_name)
        
        if not user_id:
            messagebox.showerror("Error", "Invalid user selection!")
            return
        
        cursor.execute("INSERT INTO campaigns (title, description, created_by) VALUES (%s, %s, %s)", (title, description, user_id))
        conn.commit()
        
        messagebox.showinfo("Success", "Campaign added successfully!")
        add_window.destroy()
        refresh_campaigns()
    
    add_window = tk.Toplevel(root)
    add_window.title("Add Campaign")
    add_window.geometry("300x250")
    
    tk.Label(add_window, text="Campaign Title:").pack(pady=5)
    entry_title = tk.Entry(add_window, width=30)
    entry_title.pack(pady=5)
    
    tk.Label(add_window, text="Description:").pack(pady=5)
    entry_description = tk.Entry(add_window, width=30)
    entry_description.pack(pady=5)

    tk.Label(add_window, text="Select User:").pack(pady=5)
    user_dropdown = ttk.Combobox(add_window, values=list(users_dict.keys()), state="readonly", font=("Arial", 12))
    user_dropdown.pack(pady=5)
    
    tk.Button(add_window, text="Save", command=save_campaign).pack(pady=10)

def refresh_campaigns():
    global campaigns_dict
    campaigns_dict = load_campaigns()
    campaign_dropdown['values'] = list(campaigns_dict.keys())
    campaign_dropdown.set("Select a Campaign")

def logMessages(toMobileNumber, amount, name, prename, campaign_id, user_id):
    now = datetime.now()
    query = """
    INSERT INTO messages (time_sent, date_sent, toMobileNumber, amount, full_name, campaign_id, sent_by)
    VALUES (%s, %s, %s, %s, %s, %s, %s)
    """
    values = (now.strftime('%H:%M:%S'), now.strftime('%Y-%m-%d'), toMobileNumber, amount, f"{prename} {name}", campaign_id, user_id)
    cursor.execute(query, values)
    conn.commit()
    print(f"Log added for {name} ({toMobileNumber}) in messages table with campaign ID {campaign_id} and user ID {user_id}.")

def select_file():
    global file_path
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsm;*.xlsx;*.xls")],
        title="Select an Excel File"
    )
    
    if file_path:
        lbl_file.config(text=f"Selected File: {file_path}")
        btn_start.config(state=tk.NORMAL)  # Enable start button



def process_file():
    global file_path

    selected_campaign = campaign_dropdown.get()
    if selected_campaign not in campaigns_dict:
        messagebox.showerror("Error", "Please select a campaign first!")
        return

    campaign_id = campaigns_dict[selected_campaign]  # Get selected campaign ID
    selected_user = user_dropdown.get()
    if not selected_user:
        messagebox.showerror("Error", "Please select a user!")
        return

    user_id = users_dict.get(selected_user)
    
    if not file_path:
        messagebox.showerror("Error", "Please select an Excel file first!")
        return

    try:
        # Load the selected Excel file
        df = pd.read_excel(file_path, engine="openpyxl")
        print("Columns in Excel file:", df.columns.tolist())

        # Rename and keep only relevant columns
        df.columns = ["Phone#", "IV", "Person in Charge", "Total OS", "Subject"]
        df = df[["Phone#", "IV", "Person in Charge", "Total OS"]]

        # Convert phone numbers to strings
        df["Phone#"] = df["Phone#"].astype(str).str.strip()

        # Initialize Selenium WebDriver
        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)  # Keep browser open
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

        # Open WhatsApp Web
        WHATSAPP_WEB_URL = "https://web.whatsapp.com/"
        driver.get(WHATSAPP_WEB_URL)

        # Wait for the QR code to load, then for the chat page to be ready (we'll wait for the search box to appear)
        WebDriverWait(driver, 120).until(
            EC.presence_of_element_located((By.XPATH, '//div[@class="_3F6oQ"]'))  # This is the search box element
        )

        # Now the QR code has been scanned and the web interface is ready to use
        print("WhatsApp Web is ready. Proceeding with sending messages.")

        # Iterate through the rows and send messages
        for index, row in df.iterrows():
            phone = row["Phone#"]
            prename = row["IV"]
            name = row["Person in Charge"]
            amount = row["Total OS"]

###############################################################
            message = """\
            Dear {name},
            This is a reminder for the {prename} {amount} OS.
        """
###############################################################
            
            encoded_message = urllib.parse.quote(message)

            # Open WhatsApp chat by phone number
            whatsapp_url = f"https://web.whatsapp.com/send?phone={phone}&text={encoded_message}"
            driver.get(whatsapp_url)

            # Wait for the input field to be present (chat has opened)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="1"]'))
            )

            # Find the message input field
            message_input = driver.find_element(By.XPATH, '//div[@contenteditable="true"][@data-tab="1"]')

            # Send the message
            message_input.send_keys(message)

            # Find and click the send button
            send_button = driver.find_element(By.XPATH, '//button[@aria-label="Send"]')
            send_button.click()

            print(f"Message sent to {name} ({phone})")
            logMessages(phone, amount, name, prename, campaign_id, user_id)

            time.sleep(5)  # Pause between each message

        # Close browser
        driver.quit()
        messagebox.showinfo("Success", "Messages sent successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to process file: {e}")

        

# Create Tkinter window
root = tk.Tk()
root.title("Arope WhatsApp Automation")
root.geometry("600x400")
root.resizable(True, True)
root.iconbitmap("C:/Users/VB/Desktop/Whatsapp Project/Nav_bullet.ico")

# Fetch users and store in dictionary
users_dict = load_users()

# Fetch campaigns and store in dictionary
campaigns_dict = load_campaigns()


# UI Elements for Add Campaign Window
lbl_campaign = tk.Label(root, text="Select a Campaign", font=("Arial", 15))
lbl_campaign.pack(pady=5)

# Dropdown for campaigns
campaign_dropdown = ttk.Combobox(root, values=list(campaigns_dict.keys()), state="readonly", font=("Arial", 12))
campaign_dropdown.pack(pady=5)
campaign_dropdown.set("Select a Campaign")  # Default text
tk.Button(root, text="Add Campaign", command=add_campaign).pack(side=tk.TOP)

# Horizontal line separating File Selection from User Selection
canvas2 = tk.Canvas(root, height=2, bg="gray", bd=0, highlightthickness=0)
canvas2.pack(fill="x", pady=10)

# UI Elements for Start Process Window
lbl_instruction = tk.Label(root, text="Select an Excel file", font=("Arial", 15))
lbl_instruction.pack(pady=10)

btn_select = tk.Button(root, text="Select", command=select_file, font=("Arial", 12))
btn_select.pack(pady=5)

lbl_file = tk.Label(root, text="No file selected", font=("Arial", 12))
lbl_file.pack(pady=2)


canvas2 = tk.Canvas(root, height=2, bg="gray", bd=0, highlightthickness=0)
canvas2.pack(fill="x", pady=10)

tk.Label(root, text="Select User", font=("Arial", 15)).pack(pady=5)
user_dropdown = ttk.Combobox(root, values=list(users_dict.keys()), state="readonly", font=("Arial", 12))
user_dropdown.pack(pady=5)
user_dropdown.set("Select a User")  # Default text

btn_start = tk.Button(root, text="Start Process", command=process_file, font=("Arial", 12), state=tk.DISABLED)
btn_start.pack(pady=20)

# Run the GUI
root.mainloop()

# Close database connection when the script exits
cursor.close()
conn.close()
print("Database connection closed.")

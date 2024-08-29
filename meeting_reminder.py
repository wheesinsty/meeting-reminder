"""
meeting_reminder.py

This script checks the excel sheet for meetings scheduled for tomorrow and sends a reminder email.
"""
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright
import pandas as pd
import numpy as np

# open the excel sheet
while True:
    INPUT_FILEPATH = input("Please enter the filepath to the excel sheet including the file format (.csv or .xlsx): ")
    if ".csv" in INPUT_FILEPATH:
        try:
            df = pd.read_csv(INPUT_FILEPATH)
        except:
            print("Incorrect filename or format. Please double check the filepath and filetype.")
            continue
    elif ".xlsx" in INPUT_FILEPATH:
        try:
            df = pd.read_excel(INPUT_FILEPATH)
        except:
            print("Incorrect filename or format. Please double check the filepath and filetype.")
            continue
    break

for row in range(len(df["To"])):
    
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        default_context = browser.contexts[0]
        page = default_context.pages[0]
        
        client = df.loc[row, "To"]
        now = datetime.now()
        tomorrow = (now + timedelta(days=1)).date()
        meeting_date = df.loc[row, "Meeting date"]
        start_time = str(df.loc[row, "Start time"])
        
        # deletes past meetings from the excel sheet
        if meeting_date < datetime.now():
            df.loc[row] = np.nan
            continue
    
        # checks if the meeting is scheduled for tomorrow and sends a reminder email using Outlook
        if tomorrow == meeting_date:
            page.goto("https://outlook.office.com/mail/")
            page.wait_for_timeout(3000)
            page.get_by_label("New mail").first.click()
            page.wait_for_timeout(1000)
            page.get_by_label("To", exact=True).fill(client)
            page.get_by_placeholder("Add a subject").fill("Test email")
            page.get_by_label("Message body, press Alt+F10").fill("You have a meeting scheduled at " + str(meeting_date) + " " + start_time + ".")
            page.get_by_title("Send (Ctrl+Enter)").click()
        
    browser.close()
    df.loc[row, "Status"] = "Reminder sent on " + str(now)

# delete all empty rows
df.dropna(how='all')
print(df)

# save the updated excel sheet
if ".csv" in INPUT_FILEPATH:
    df.to_csv(INPUT_FILEPATH, index=False)
else:
    df.to_excel(INPUT_FILEPATH, index=False)
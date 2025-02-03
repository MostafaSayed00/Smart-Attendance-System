import time
import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
from mfrc522 import SimpleMFRC522
import RPi.GPIO as GPIO
import os

# Paths
EXCEL_PATH = '/home/pi/Desktop/Last_code/rfid_cards.xlsx'

# Gmail credentials
GMAIL_USER = 'madina.higher.institute@gmail.com'
GMAIL_PASSWORD = 'Madina.higher.institute2000'
TO_EMAIL = 'nazzazzz@gmail.com'

# Setup RFID reader
reader = SimpleMFRC522()

# Ensure GPIO is cleaned up in case of previous use
try:
    GPIO.cleanup()  # Reset GPIO settings to avoid conflicts
except RuntimeError:
    pass  # Ignore if cleanup fails

# Set GPIO mode
GPIO.setmode(GPIO.BCM)

# LED and buzzer pins
GREEN_LED = 20
RED_LED = 21
BUZZER = 18
GPIO.setup(GREEN_LED, GPIO.OUT)
GPIO.setup(RED_LED, GPIO.OUT)
GPIO.setup(BUZZER, GPIO.OUT)

# Read and initialize data
try:
    data = pd.read_excel(EXCEL_PATH)
except Exception as e:
    print(f"Failed to read Excel file: {e}")
    GPIO.cleanup()
    exit()

# Check if sheet is not empty before proceeding
if data.empty:
    print("Excel sheet is empty. Please check the file.")
    GPIO.cleanup()
    exit()

# Print all registered card UIDs for debugging purposes
try:
    registered_uids = data.iloc[:, 0].astype(str).values  # Ensure UID is treated as string
    print("Registered Card UIDs:", registered_uids)
except Exception as e:
    print(f"Failed to read registered card UIDs: {e}")

# Find the next empty column for attendance marking
next_column = data.shape[1]
date_now = datetime.datetime.now().strftime('%Y-%m-%d')

# Add a new column if necessary
if next_column >= data.shape[1]:
    data[date_now] = pd.NA  # Add a new column with today's date

attended_cards = set()

print("Starting attendance system...")
start_time = time.time()

# Attendance loop
while True:
    try:
        current_time = time.time()
        elapsed_time = current_time - start_time

        if elapsed_time > 240:  # Attendance period ends after 4 minutes
            print("Attendance finished.")
            GPIO.output(RED_LED, GPIO.HIGH)
            time.sleep(2)
            GPIO.output(RED_LED, GPIO.LOW)
            break

        print("Waiting for card...")
        uid, _ = reader.read()
        str_uid = str(uid)
        print(f"Card UID Read: {str_uid}")

        # Check if UID is registered
        if str_uid not in registered_uids:
            print("This card is not registered.")
            GPIO.output(RED_LED, GPIO.HIGH)
            time.sleep(1)
            GPIO.output(RED_LED, GPIO.LOW)
            continue

        # Check for duplicate attendance
        if str_uid in attended_cards:
            print("This card already took attendance before.")
            GPIO.output(RED_LED, GPIO.HIGH)
            time.sleep(1)
            GPIO.output(RED_LED, GPIO.LOW)
            continue

        # Mark attendance
        attended_cards.add(str_uid)
        student_index = data[data.iloc[:, 0].astype(str) == str_uid].index[0]
        student_name = data.iloc[student_index, 2]
        current_time_str = datetime.datetime.now().strftime('%H:%M:%S')

        if elapsed_time <= 120:  # First 2 minutes
            print(f"Hello {student_name}, have a good day! Time: {current_time_str}")
            data.loc[student_index, date_now] = "\u2714"  # Checkmark
            GPIO.output(GREEN_LED, GPIO.HIGH)
        elif elapsed_time <= 240:  # Between 2 and 4 minutes
            print(f"Hello {student_name}, you are late. Time: {current_time_str}")
            data.loc[student_index, date_now] = "\u2718"  # Crossmark
            GPIO.output(RED_LED, GPIO.HIGH)

        time.sleep(1)
        GPIO.output(GREEN_LED, GPIO.LOW)
        GPIO.output(RED_LED, GPIO.LOW)

    except KeyboardInterrupt:
        print("Exiting attendance system.")
        break
    except Exception as e:
        print(f"Error: {e}")
        continue

# Mark "not attend" for missing students
for i in range(len(data)):
    if pd.isnull(data.loc[i, date_now]):
        data.loc[i, date_now] = "not attend"
        print(f"Marked 'not attend' for {data.iloc[i, 2]}.")

# Save updated data
try:
    data.to_excel(EXCEL_PATH, index=False)
    print("Attendance saved to Excel.")
except Exception as e:
    print(f"Failed to save Excel: {e}")

# Send the Excel sheet via email using mutt
try:
    os.system(f'echo "Here is the Excel sheet you requested." | mutt -s "Your Excel Sheet" -a {EXCEL_PATH} -- {TO_EMAIL}')
    print("Excel sheet emailed successfully.")
except Exception as e:
    print(f"Failed to send email: {e}")

# Cleanup GPIO
print("Cleaning up GPIO...")
GPIO.cleanup()

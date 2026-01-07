import pandas as pd
from bs4 import BeautifulSoup
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# --------------------- CONFIGURATION ---------------------
# GitHub Secrets se ID/Password uthayenge (Security ke liye)
YOUR_EMAIL = os.environ.get('EMAIL_USER')
APP_PASSWORD = os.environ.get('EMAIL_PASS')

# Jisko email bhejna hai uska address yahan likh dein
TO_EMAIL = "priyansh.khandelwal@rosierfoods.com" 

EXCEL_FILE = "blinkit_rosier_products.xlsx"
HTML_FILE = "blinkit.html"

# Check credentials before starting
if not YOUR_EMAIL or not APP_PASSWORD:
    print("Error: GitHub Secrets (EMAIL_USER / EMAIL_PASS) set nahi hain!")
    exit()

# ------------------------------------------------------

# 1. HTML Read karna
print("Reading HTML file...")
try:
    with open(HTML_FILE, 'r', encoding='utf-8') as file:
        html_content = file.read()
except FileNotFoundError:
    print(f"Error: '{HTML_FILE}' file Repository me nahi mili! Please upload karein.")
    exit()

soup = BeautifulSoup(html_content, 'html.parser')

# Update: 'role': 'button' generic ho sakta hai, blinkit classes change karta rehta hai.
# Agar aapka purana logic kaam kar raha tha to wahi rakha hai.
product_containers = soup.find_all('div', attrs={'role': 'button', 'tabindex': '0'})

print(f"Total containers found: {len(product_containers)}")

product_details = []

for container in product_containers:
    # Title
    title_div = container.find('div', class_=re.compile(r'tw-text-300.*tw-font-semibold.*tw-line-clamp-2'))
    if not title_div:
        continue
   
    product_name = title_div.get_text(strip=True)
   
    # Filter: Sirf Rosier products
    if 'rosier' not in product_name.lower():
        continue
   
    # ====================== VARIANT / QUANTITY निकालना ======================
    variant = "-"
   
    next_div = title_div.find_next_sibling('div')
    if next_div:
        text = next_div.get_text(strip=True)
        if any(unit in text.lower() for unit in ['kg', 'g', 'ml', 'l', 'pack', 'piece', 'pcs']):
            variant = text
   
    if variant == "-":
        container_text = container.get_text(separator=" ", strip=True)
        match = re.search(r'(\d+(?:\.\d+)?\s*(?:kg|g|ml|l|piece|pack|pcs|bottle|jar|box|kgx|gx|ltr|litre|kilogram))\b',
                         container_text, re.IGNORECASE)
        if match:
            variant = match.group(1).strip()
   
    if variant == "-" and any(unit in product_name.lower() for unit in ['kg', 'g', 'ml', 'l']):
        words = product_name.split()
        for i in range(len(words)-1, -1, -1):
            if any(unit in words[i].lower() for unit in ['kg', 'g', 'ml', 'l', 'pack', 'piece']):
                start = max(0, i-1)
                variant = ' '.join(words[start:i+1])
                break
   
    # Stock status
    text_content = container.get_text(separator=" ", strip=True).lower()
    has_out_stock = any(phrase in text_content for phrase in [
        'out of stock', 'outofstock', 'currently unavailable', 'not available'
    ])
   
    # Price
    price = "-"
    price_div = container.find('div', class_=re.compile(r'tw-text-200.*tw-font-semibold'))
    if price_div:
        price = price_div.get_text(strip=True)
    else:
        possible_price = container.find(string=re.compile(r'^₹[0-9,]+$'))
        if possible_price:
            price = possible_price.strip()
   
    stock_status = "Out of Stock" if has_out_stock else "In Stock"
   
    product_details.append({
        "Title": product_name,
        "Variant": variant.strip() if variant != "-" else "-",
        "Price": price,
        "Stock": stock_status
    })

# 2. Excel Save karna
if product_details:
    df = pd.DataFrame(product_details)
    print(f"\nTotal Rosier Products Found: {len(df)}")
    # print(df) # Console ko clean rakhne ke liye comment kiya
   
    df.to_excel(EXCEL_FILE, index=False)
    print(f"Data saved to → {EXCEL_FILE}")
else:
    print("Koi bhi Rosier product nahi mila HTML me.")
    exit()

# --------------------- EMAIL SENDING ---------------------
print("Preparing to send email...")

msg = MIMEMultipart()
msg['From'] = YOUR_EMAIL
msg['To'] = TO_EMAIL
msg['Subject'] = 'Blinkit - Latest Rosier Products List'

body = f"""Hi Automailer,

PFA Blinkit Rosier products.
Total Products: {len(product_details)}

Regards,
Priyansh
"""
msg.attach(MIMEText(body, 'plain'))

if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={EXCEL_FILE}')
    msg.attach(part)
else:
    print("Error: Excel file create nahi hui, email attach nahi kar paaye.")
    exit()

try:
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(YOUR_EMAIL, APP_PASSWORD)
    server.send_message(msg)
    server.quit()
    print(f"Success! Email sent to {TO_EMAIL}")
except Exception as e:
    print(f"Email Failed! Error: {e}")

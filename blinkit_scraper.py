import pandas as pd
from bs4 import BeautifulSoup
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from openpyxl import load_workbook
from openpyxl.styles import Font

# --------------------- CONFIGURATION ---------------------
YOUR_EMAIL = os.environ.get('EMAIL_USER')
APP_PASSWORD = os.environ.get('EMAIL_PASS')
TO_EMAIL = "priyanshkhandelwal@rosierfoods.com" # <--- Receiver Email yaha dalein

EXCEL_FILE = "blinkit_rosier_products.xlsx"
HTML_FILE = "blinkit.html"

# Security Check
if not YOUR_EMAIL or not APP_PASSWORD:
    # Local run ke liye agar environment variable set nahi hai to error na de (for testing)
    pass 

# --------------------- SCRAPING LOGIC ---------------------
print("Reading HTML file...")
try:
    with open(HTML_FILE, 'r', encoding='utf-8') as file:
        html_content = file.read()
except FileNotFoundError:
    print(f"Error: '{HTML_FILE}' file nahi mili!")
    exit()

soup = BeautifulSoup(html_content, 'html.parser')

# Container dhundne ka logic
product_containers = soup.find_all('div', attrs={'role': 'button', 'tabindex': '0'})
print(f"Total containers found: {len(product_containers)}")

product_details = []

for container in product_containers:
    # 1. Title Extraction
    title_div = container.find('div', class_=re.compile(r'tw-text-300.*tw-font-semibold.*tw-line-clamp-2'))
    if not title_div:
        continue
    
    product_name = title_div.get_text(strip=True)
    
    # Filter: Only Rosier
    if 'rosier' not in product_name.lower():
        continue

    # ============================================================
    # 2. LINK EXTRACTION (UPDATED - AGGRESSIVE SEARCH)
    # ============================================================
    product_url = None
    
    # Tarika A: Kya container ke andar koi 'a' tag hai?
    link_tag = container.find('a', href=True)
    
    # Tarika B: Agar andar nahi, to kya container khud kisi 'a' tag ke andar hai? (Parent)
    if not link_tag:
        link_tag = container.find_parent('a', href=True)
        
    # Tarika C: Title div ke aas paas dhundo
    if not link_tag:
        link_tag = title_div.find_parent('a', href=True)

    # Agar link mil gaya to URL banayein
    if link_tag:
        href_val = link_tag['href']
        if href_val.startswith('/'):
            product_url = "https://blinkit.com" + href_val
        else:
            product_url = href_val
    else:
        # Debugging ke liye: Pata chale kyu nahi mila
        print(f"Warning: Link nahi mila for -> {product_name}")

    # ============================================================

    # 3. Variant Extraction
    variant = "-"
    next_div = title_div.find_next_sibling('div')
    if next_div:
        text = next_div.get_text(strip=True)
        if any(unit in text.lower() for unit in ['kg', 'g', 'ml', 'l', 'pack', 'piece', 'pcs']):
            variant = text
    
    if variant == "-":
        container_text = container.get_text(separator=" ", strip=True)
        match = re.search(r'(\d+(?:\.\d+)?\s*(?:kg|g|ml|l|piece|pack|pcs|bottle|jar|box|kgx|gx|ltr|litre|kilogram))\b', container_text, re.IGNORECASE)
        if match:
            variant = match.group(1).strip()

    if variant == "-" and any(unit in product_name.lower() for unit in ['kg', 'g', 'ml', 'l']):
        words = product_name.split()
        for i in range(len(words)-1, -1, -1):
            if any(unit in words[i].lower() for unit in ['kg', 'g', 'ml', 'l', 'pack', 'piece']):
                start = max(0, i-1)
                variant = ' '.join(words[start:i+1])
                break

    # 4. Stock & Price
    text_content = container.get_text(separator=" ", strip=True).lower()
    has_out_stock = any(phrase in text_content for phrase in ['out of stock', 'outofstock', 'currently unavailable'])
    stock_status = "Out of Stock" if has_out_stock else "In Stock"

    price = "-"
    price_div = container.find('div', class_=re.compile(r'tw-text-200.*tw-font-semibold'))
    if price_div:
        price = price_div.get_text(strip=True)
    else:
        possible_price = container.find(string=re.compile(r'^₹[0-9,]+$'))
        if possible_price:
            price = possible_price.strip()

    product_details.append({
        "Title": product_name,
        "Variant": variant.strip(),
        "Price": price,
        "Stock": stock_status,
        "Hidden_URL": product_url
    })

# --------------------- EXCEL GENERATION ---------------------
if product_details:
    df = pd.DataFrame(product_details)
    
    # Excel Save
    df.to_excel(EXCEL_FILE, index=False)
    
    # Hyperlink Styling
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        title_cell = row[0]       # Column A (Title)
        url_cell = row[4]         # Column E (Hidden_URL)
        
        # Check karein ki URL cell khali to nahi hai
        if url_cell.value:
            title_cell.hyperlink = url_cell.value
            title_cell.font = Font(color="0000FF", underline="single")
        else:
            # Agar URL nahi mila to Red color kar do taaki pata chale
            title_cell.font = Font(color="FF0000") 

    ws.delete_cols(5) # URL column delete
    wb.save(EXCEL_FILE)
    print(f"Data saved with Links to → {EXCEL_FILE}")
else:
    print("No Rosier products found.")
    exit()

# --------------------- EMAIL SENDING ---------------------
if not YOUR_EMAIL or not APP_PASSWORD:
    print("Skipping Email (No Credentials). File saved locally.")
    exit()

print("Sending Email...")
msg = MIMEMultipart()
msg['From'] = YOUR_EMAIL
msg['To'] = TO_EMAIL
msg['Subject'] = 'Blinkit - Latest Rosier Products'

body = f"Hi Automailer PFA Rosier blinkit products.\nTotal: {len(product_details)}"
msg.attach(MIMEText(body, 'plain'))

with open(EXCEL_FILE, 'rb') as attachment:
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', f'attachment; filename={EXCEL_FILE}')
msg.attach(part)

try:
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(YOUR_EMAIL, APP_PASSWORD)
    server.send_message(msg)
    server.quit()
    print("Email Sent!")
except Exception as e:
    print(f"Email Failed: {e}")

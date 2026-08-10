import json
import os
import re
import time
import openpyxl
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


def get_excel_path():
    """
    Resolves file path dynamically for GitHub Actions or local execution.
    """
    primary_path = "FK Walkthrough 10th aug.xlsx"
    local_path = r"C:\Users\sai\OneDrive\Desktop\flipkart\FK Walkthrough 10th aug.xlsx"

    if os.path.exists(primary_path):
        return primary_path
    elif os.path.exists(local_path):
        return local_path
    elif os.path.exists("FK Walkthrough 10th aug"):
        return "FK Walkthrough 10th aug"
    else:
        raise FileNotFoundError("Could not locate 'FK Walkthrough 10th aug.xlsx'.")


def get_clean_price(html_source):
    """
    Extracts numerical price from Flipkart product page HTML.
    """
    soup = BeautifulSoup(html_source, "html.parser")

    # 1. Primary Method: Inspect application/ld+json structured data
    json_ld_scripts = soup.find_all("script", type="application/ld+json")
    for script in json_ld_scripts:
        if script.string:
            try:
                data = json.loads(script.string)
                if isinstance(data, list):
                    data = data[0]
                if "offers" in data:
                    offers = data["offers"]
                    if isinstance(offers, list):
                        offers = offers[0]
                    price = offers.get("price") or offers.get("lowPrice")
                    if price:
                        return str(price).replace(",", "").strip()
            except Exception:
                continue

    # 2. Fallback Method: Target main selling price using robust regex
    matches = re.findall(r"₹\s*([0-9,]+)", html_source)
    if matches:
        return matches[0].replace(",", "").strip()

    return "Not Found"


def main():
    excel_file = get_excel_path()
    wb = openpyxl.load_workbook(excel_file)

    # Configure Headless Chrome for GitHub Actions Cloud Execution
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    )

    print("Launching Headless Chrome Driver for Flipkart...")
    driver = webdriver.Chrome(options=chrome_options)

    # Sheets to skip (case-insensitive check)
    skip_sheets = {"amzon home", "amazon home", "sheet1", "sheet 1"}

    # Loop through every sheet in the workbook
    for sheet_name in wb.sheetnames:
        if sheet_name.strip().lower() in skip_sheets:
            print(f"\nSkipping sheet: '{sheet_name}'")
            continue

        ws = wb[sheet_name]
        print(f"\n{'=' * 50}\nProcessing sheet: '{sheet_name}'\n{'=' * 50}")

        # Get header names from the first row
        headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]]

        # Dynamically find column indices
        url_col_idx = headers.index("URL") + 1 if "URL" in headers else 6
        
        if "Current Price" in headers:
            price_col_idx = headers.index("Current Price") + 1
        elif "Current" in headers:
            price_col_idx = headers.index("Current") + 1
        else:
            print(f"Skipping sheet '{sheet_name}': 'Current Price' column not found.")
            continue

        stock_col_idx = headers.index("FK Stock") + 1 if "FK Stock" in headers else 13

        total_rows = ws.max_row

        # Loop through rows (starting from row 2 to skip headers)
        for row in range(2, total_rows + 1):
            # Check FK Stock condition (must be >= 50)
            stock_val = ws.cell(row=row, column=stock_col_idx).value
            try:
                stock_count = float(stock_val) if stock_val is not None else 0
            except (ValueError, TypeError):
                stock_count = 0

            if stock_count < 50:
                print(f"[{row-1}/{total_rows-1}] Stock is {stock_val} (< 50). Skipping Row {row}.")
                continue

            url_cell = ws.cell(row=row, column=url_col_idx).value
            url = str(url_cell).strip() if url_cell else ""

            if not url.startswith("http"):
                print(f"[{row-1}/{total_rows-1}] Invalid or missing URL. Skipping Row {row}.")
                continue

            try:
                driver.get(url)
                time.sleep(2.5)  # Wait for dynamic rendering

                price = get_clean_price(driver.page_source)

                # Write price directly into the 'Current Price' column
                if price != "Not Found":
                    try:
                        price_val = float(price) if "." in price else int(price)
                    except ValueError:
                        price_val = price

                    ws.cell(row=row, column=price_col_idx, value=price_val)
                    price_str = f"₹{price}"
                else:
                    ws.cell(row=row, column=price_col_idx, value="Not Found")
                    price_str = "Not Found"

                print(f"[{row-1}/{total_rows-1}] Row {row} Updated -> Price: {price_str}")

            except Exception as err:
                print(f"[{row-1}/{total_rows-1}] Row {row} Error processing URL: {err}")

    driver.quit()

    # Save changes back to the original file
    wb.save(excel_file)
    print("\n" + "=" * 50)
    print(f"All prices successfully updated across all eligible sheets in '{excel_file}'!")


if __name__ == "__main__":
    main()

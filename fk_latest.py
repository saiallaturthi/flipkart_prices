import json
import os
import re
import time
import openpyxl
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


def get_excel_path():
    primary_path = r"C:\Users\sai\OneDrive\Desktop\flipkart\FK Walkthrough 10th aug.xlsx"
    secondary_path = r"C:\Users\sai\OneDrive\Desktop\flipkart\FK Walkthrough 10th aug"

    if os.path.exists(primary_path):
        return primary_path
    elif os.path.exists(secondary_path):
        return secondary_path
    else:
        raise FileNotFoundError(
            f"Could not locate 'FK Walkthrough 10th aug.xlsx' in {os.path.dirname(primary_path)}"
        )


def get_clean_price(html_source):
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

    # Attach to existing Chrome session
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

    try:
        driver = webdriver.Chrome(options=chrome_options)
    except Exception as e:
        print(f"Error connecting to Chrome on port 9222: {e}")
        return

    # Issues output folder setup
    issues_dir = r"C:\Users\sai\OneDrive\Desktop\flipkart\Issues"
    os.makedirs(issues_dir, exist_ok=True)

    skip_sheets = {"amzon home", "sheet1", "sheet 1"}

    for sheet_name in wb.sheetnames:
        if sheet_name.strip().lower() in skip_sheets:
            print(f"\nSkipping sheet: '{sheet_name}'")
            continue

        ws = wb[sheet_name]
        print(f"\n{'=' * 50}\nProcessing sheet: '{sheet_name}'\n{'=' * 50}")

        headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]]

        def get_col_idx(col_name, default_idx=-1):
            return headers.index(col_name) + 1 if col_name in headers else default_idx

        # Locate dynamically required columns
        url_col_idx = get_col_idx("URL", 6)
        price_col_idx = get_col_idx("Current Price", -1)
        stock_col_idx = get_col_idx("FK Stock", 13)
        az_price_col_idx = get_col_idx("AZ Price", -1)
        diff_col_idx = get_col_idx("Price Difference", -1)

        # Output issue columns matching target schema
        fsn_col_idx = get_col_idx("Flipkart Serial Number", -1)
        sku_col_idx = get_col_idx("Seller SKU Id", -1)
        title_col_idx = get_col_idx("Product Title", -1)
        remark_col_idx = get_col_idx("Remark", -1)

        if price_col_idx == -1:
            print(f"Skipping sheet '{sheet_name}': 'Current Price' column not found.")
            continue

        total_rows = ws.max_row
        sheet_issues = []

        # Step 1: Scrape & Update Prices for Stock >= 50
        for row in range(2, total_rows + 1):
            stock_val = ws.cell(row=row, column=stock_col_idx).value
            try:
                stock_count = float(stock_val) if stock_val is not None else 0
            except (ValueError, TypeError):
                stock_count = 0

            # Condition: Stock must be >= 50
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
                time.sleep(1.5)

                price = get_clean_price(driver.page_source)

                if price != "Not Found":
                    try:
                        price_val = float(price) if "." in price else int(price)
                    except ValueError:
                        price_val = price

                    ws.cell(row=row, column=price_col_idx, value=price_val)
                else:
                    ws.cell(row=row, column=price_col_idx, value="Not Found")
                    price_val = None

                print(f"[{row-1}/{total_rows-1}] Row {row} Updated -> Price: {price}")

                # Step 2: Calculate difference & check for Issue (>10 or <-10)
                if price_val is not None and isinstance(price_val, (int, float)):
                    az_val = ws.cell(row=row, column=az_price_col_idx).value if az_price_col_idx != -1 else None
                    diff_val = ws.cell(row=row, column=diff_col_idx).value if diff_col_idx != -1 else None

                    try:
                        az_num = float(az_val) if az_val is not None else None
                    except (ValueError, TypeError):
                        az_num = None

                    try:
                        diff_num = float(diff_val) if diff_val is not None else None
                    except (ValueError, TypeError):
                        diff_num = None

                    # If Price Difference is missing, calculate it directly (AZ Price - Current Price)
                    if diff_num is None and az_num is not None:
                        diff_num = az_num - price_val

                    # Filter condition: Price Difference strictly > 10 or < -10
                    if diff_num is not None and abs(diff_num) > 10:
                        fsn = ws.cell(row=row, column=fsn_col_idx).value if fsn_col_idx != -1 else ""
                        sku = ws.cell(row=row, column=sku_col_idx).value if sku_col_idx != -1 else ""
                        title = ws.cell(row=row, column=title_col_idx).value if title_col_idx != -1 else ""
                        remark = ws.cell(row=row, column=remark_col_idx).value if remark_col_idx != -1 else ""

                        # Generate Remark automatically if blank
                        if not remark and az_num is not None:
                            if price_val > az_num:
                                remark = "Current price is More than Event price"
                            elif price_val < az_num:
                                remark = "Current price is Less than Event price"

                        sheet_issues.append([
                            fsn,
                            sku,
                            title,
                            az_num if az_num is not None else az_val,
                            price_val,
                            remark
                        ])

            except Exception as err:
                print(f"[{row-1}/{total_rows-1}] Row {row} Error processing URL: {err}")

        # Step 3: Export per-sheet Issues file if differences exist
        if sheet_issues:
            issues_filename = f"{sheet_name.strip()} issues.xlsx"
            issues_path = os.path.join(issues_dir, issues_filename)

            issues_wb = openpyxl.Workbook()
            issues_ws = issues_wb.active
            issues_ws.title = f"{sheet_name.strip()} Issues"

            # Image 2 Headers
            issues_ws.append([
                "Flipkart Serial Number",
                "Seller SKU Id",
                "Product Title",
                "AZ Price",
                "Current Price",
                "Remark"
            ])

            for issue_row in sheet_issues:
                issues_ws.append(issue_row)

            issues_wb.save(issues_path)
            print(f"--> Saved issue file: '{issues_path}' ({len(sheet_issues)} rows)")

    # Save original Excel workbook with updated prices
    wb.save(excel_file)
    print("\n" + "=" * 50)
    print(f"All prices and issues files successfully processed for '{excel_file}'!")


if __name__ == "__main__":
    main()
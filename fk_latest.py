import json
import os
import re
import time
import openpyxl
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


def get_excel_path(base_name):
    """
    Finds file path dynamically whether the filename has .xlsx or no extension.
    """
    possible_paths = [
        base_name,
        f"{base_name}.xlsx",
        os.path.join(r"C:\Users\sai\OneDrive\Desktop\flipkart", base_name),
        os.path.join(r"C:\Users\sai\OneDrive\Desktop\flipkart", f"{base_name}.xlsx"),
    ]

    for path in possible_paths:
        if os.path.exists(path):
            return path

    return None


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

    # 2. Fallback Method: Target main selling price using regex
    matches = re.findall(r"₹\s*([0-9,]+)", html_source)
    if matches:
        return matches[0].replace(",", "").strip()

    return "Not Found"


def append_issue_row(target_file_name, row_data):
    """
    Appends a formatted discrepancy row into the target issue workbook.
    Expected row_data format:
    [Flipkart Serial Number, Seller SKU Id, Product Title, AZ Price, Current Price, Remark]
    """
    target_path = get_excel_path(target_file_name)
    if not target_path:
        print(f" [Warning] Could not locate issue workbook for '{target_file_name}'. Skipping dump.")
        return

    try:
        issue_wb = openpyxl.load_workbook(target_path)
        issue_ws = issue_wb.active

        # If empty file, initialize headers matching Image 2
        if issue_ws.max_row == 1 and issue_ws.cell(row=1, column=1).value is None:
            issue_ws.append([
                "Flipkart Serial Number",
                "Seller SKU Id",
                "Product Title",
                "AZ Price",
                "Current Price",
                "Remark"
            ])

        issue_ws.append(row_data)
        issue_wb.save(target_path)
    except Exception as e:
        print(f" [Error] Failed to write to {target_file_name}: {e}")


def main():
    main_excel_file = get_excel_path("FK Walkthrough 10th aug")
    if not main_excel_file:
        raise FileNotFoundError("Could not locate main file 'FK Walkthrough 10th aug.xlsx'.")

    wb = openpyxl.load_workbook(main_excel_file)

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

    # Sheets to skip
    skip_sheets = {"amzon home", "amazon home", "sheet1", "sheet 1"}

    for sheet_name in wb.sheetnames:
        clean_sheet_name = sheet_name.strip()
        if clean_sheet_name.lower() in skip_sheets:
            print(f"\nSkipping sheet: '{sheet_name}'")
            continue

        ws = wb[sheet_name]
        print(f"\n{'=' * 60}\nProcessing sheet: '{sheet_name}'\n{'=' * 60}")

        headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]]

        def get_col_idx(name, default_idx=-1):
            return headers.index(name) + 1 if name in headers else default_idx

        # Locate core column indices dynamically
        url_col_idx = get_col_idx("URL", 6)
        price_col_idx = get_col_idx("Current Price", get_col_idx("Current", -1))
        stock_col_idx = get_col_idx("FK Stock", 13)
        az_price_col_idx = get_col_idx("AZ Price", get_col_idx("Amazon Price", -1))
        fsn_col_idx = get_col_idx("Flipkart Serial Number", get_col_idx("FSN", 1))
        sku_col_idx = get_col_idx("Seller SKU Id", get_col_idx("SKU", 2))
        title_col_idx = get_col_idx("Product Title", get_col_idx("Title", 3))
        category_col_idx = get_col_idx("Category", get_col_idx("Sub-Category", -1))

        if price_col_idx == -1:
            print(f"Skipping sheet '{sheet_name}': 'Current Price' column not found.")
            continue

        total_rows = ws.max_row

        for row in range(2, total_rows + 1):
            # 1. Stock Check: Must be >= 50
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
                print(f"[{row-1}/{total_rows-1}] Invalid/Missing URL. Skipping Row {row}.")
                continue

            try:
                driver.get(url)
                time.sleep(2.5)  # Allow page to render

                current_price_str = get_clean_price(driver.page_source)

                if current_price_str != "Not Found":
                    try:
                        current_price_val = float(current_price_str) if "." in current_price_str else int(current_price_str)
                    except ValueError:
                        current_price_val = current_price_str

                    ws.cell(row=row, column=price_col_idx, value=current_price_val)
                    print_price = f"₹{current_price_val}"
                else:
                    current_price_val = None
                    ws.cell(row=row, column=price_col_idx, value="Not Found")
                    print_price = "Not Found"

                # 2. Check Price Difference against AZ Price
                az_price_val = ws.cell(row=row, column=az_price_col_idx).value if az_price_col_idx != -1 else None

                if current_price_val is not None and az_price_val is not None:
                    try:
                        az_num = float(az_price_val)
                        curr_num = float(current_price_val)
                        price_diff = abs(az_num - curr_num)

                        # Trigger issue report if difference > 10 (or non-zero difference)
                        if price_diff > 10:
                            if curr_num < az_num:
                                remark = "Current price is Less than Amazon price"
                            else:
                                remark = "Current price is More than Amazon price"

                            # Read row details for report payload
                            fsn = ws.cell(row=row, column=fsn_col_idx).value or ""
                            sku = ws.cell(row=row, column=sku_col_idx).value or ""
                            title = ws.cell(row=row, column=title_col_idx).value or ""
                            category = str(ws.cell(row=row, column=category_col_idx).value or "").lower()

                            report_row = [fsn, sku, title, az_num, curr_num, remark]

                            # Determine destination target file
                            target_file = clean_sheet_name

                            # Categorization split for Roma / Power Banks
                            if "roma" in clean_sheet_name.lower():
                                if "powerbank" in category or "power bank" in category or "power bank" in str(title).lower():
                                    target_file = "Power Banks"
                                else:
                                    target_file = "Roma"

                            append_issue_row(target_file, report_row)
                            print(f"[{row-1}/{total_rows-1}] Row {row} Updated -> Price: {print_price} | Issue logged to '{target_file}': {remark}")
                        else:
                            print(f"[{row-1}/{total_rows-1}] Row {row} Updated -> Price: {print_price} | Difference <= 10. No issue logged.")

                    except (ValueError, TypeError) as e:
                        print(f"[{row-1}/{total_rows-1}] Row {row} Updated -> Price: {print_price} | Could not calculate AZ difference: {e}")
                else:
                    print(f"[{row-1}/{total_rows-1}] Row {row} Updated -> Price: {print_price}")

            except Exception as err:
                print(f"[{row-1}/{total_rows-1}] Row {row} Error processing URL: {err}")

    driver.quit()

    wb.save(main_excel_file)
    print("\n" + "=" * 60)
    print("All prices updated and issue reports successfully created/dumped!")


if __name__ == "__main__":
    main()

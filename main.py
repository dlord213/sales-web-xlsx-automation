from concurrent.futures import ThreadPoolExecutor, as_completed
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import xlsxwriter
from data.counties_id import counties
from data.small_counties import small_counties
from bs4 import BeautifulSoup
from datetime import date

import requests
import re
import pandas as pd


def extractPropertyId(td):
    """Extract Property ID from either href or onclick attributes."""
    link = td.find("a")

    if link and link.get("href"):
        # Case 1: Direct href link containing PropertyId
        match = re.search(r"\d+$", link["href"])
        return match.group() if match else None

    elif td.get("onclick"):
        # Case 2: PropertyId inside onclick event
        match = re.search(r"PropertyId=(\d+)", td["onclick"])
        return match.group(1) if match else None

    return None  # No match found


def getSalesListingData(id: int):
    if not id:
        return None

    url = f"https://salesweb.civilview.com/Sales/SalesSearch?countyId={id}"
    headers = {"User-Agent": "Mozilla/5.0"}
    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        return None

    soup = BeautifulSoup(response.text, "html.parser")

    # Get county name from predefined list
    countyName = next(
        (county["name"] for county in counties if county["countyId"] == id), ""
    )

    # Locate the table
    table = soup.find("table", class_="table table-striped")
    if not table:
        return None  # No table found

    rows = table.find_all("tr")

    # Extract column names from <thead> or the first row
    if table.find("thead"):
        column_names = [th.text.strip() for th in table.find("thead").find_all("th")]
    else:
        column_names = [td.text.strip() for td in rows[0].find_all("td")]

    # Extract table data
    table_data = []
    for row in rows:
        cells = row.find_all("td")
        if not cells:
            continue  # Skip empty rows

        property_id = extractPropertyId(
            cells[0]
        )  # Extract Property ID from the first column
        row_data = [td.text.strip() for td in cells]

        if property_id:
            row_data.pop(0)  # Remove original Property ID
            row_data.insert(0, property_id)  # Ensure Property ID is always first

        table_data.append(row_data)

    return {
        "county": countyName,
        "countyId": id,
        "column_names": column_names,
        "data": table_data,
    }


def getPropertyDetailsByCounty(county_id: int, headless=True):
    listings = getSalesListingData(county_id)

    # Ensure listings is a dictionary before calling .get()
    if not isinstance(listings, dict) or "data" not in listings or not listings["data"]:
        county_name = (
            listings.get("county", "Unknown County")
            if isinstance(listings, dict)
            else "Unknown County"
        )
        print(f"{county_name} doesn't have any listings.")
        return (
            []
        )  # Return an empty list instead of None to prevent errors in calling functions

    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless")  # Run in headless mode
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--blink-settings=imagesEnabled=false")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=options
    )

    try:
        url = f"https://salesweb.civilview.com/Sales/SalesSearch?countyId={county_id}"
        driver.get(url)

        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "table-striped")))

        properties_data = []

        # Find "Details" buttons
        details_buttons = driver.find_elements(
            By.XPATH, "//a[contains(text(), 'Details')]"
        )

        if not details_buttons:
            details_buttons = driver.find_elements(
                By.XPATH,
                "//div[contains(@class, 'btn-primary') and contains(text(), 'Details')]",
            )

        num_properties = min(len(details_buttons), len(listings["data"]))

        for i in range(num_properties):
            try:
                # Re-locate buttons in case the page reloads
                details_buttons = driver.find_elements(
                    By.XPATH, "//a[contains(text(), 'Details')]"
                )

                if not details_buttons:
                    details_buttons = driver.find_elements(
                        By.XPATH,
                        "//div[contains(@class, 'btn-primary') and contains(text(), 'Details')]",
                    )

                wait.until(EC.element_to_be_clickable(details_buttons[i])).click()
                wait.until(
                    EC.presence_of_element_located((By.CLASS_NAME, "table-striped"))
                )

                # Extract property data
                property_data = {}
                rows = driver.find_elements(
                    By.XPATH, "(//table[@class='table table-striped'])[1]//tr"
                )

                for row in rows:
                    cols = row.find_elements(By.TAG_NAME, "td")
                    if len(cols) >= 2:
                        key = cols[0].text.strip().replace(":", "")
                        value = cols[1].text.strip()
                        property_data[key] = value

                properties_data.append(property_data)

                # Go back to the listings page
                driver.back()
                wait.until(
                    EC.presence_of_element_located((By.CLASS_NAME, "table-striped"))
                )

            except Exception as e:
                print(f"Error processing property {i+1}: {e}")

        print(
            f"Scraping complete for County ID {county_id}. Total properties: {len(properties_data)}"
        )
        return properties_data

    except Exception as e:
        print(f"Error scraping county {county_id}: {e}")
        return []

    finally:
        driver.quit()


def exportSalesListingData(id: int):
    """
    This function exports the data of every sales listing in a county depending on the id at the args.
    """
    if id:
        url = f"https://salesweb.civilview.com/Sales/SalesSearch?countyId={id}"
        headers = {"User-Agent": "Mozilla/5.0"}
        response = requests.get(url, headers=headers)

        if response.status_code == 200:
            soup = BeautifulSoup(response.text, "html.parser")

            # Extract county name
            countyName = ""
            for county in counties:
                if county["countyId"] == id:
                    countyName = county["name"]

            # Extract last updated timestamp

            # Find the table
            table = soup.find("table", class_="table table-striped")
            if table:
                # Extract all rows
                rows = table.find_all("tr")

                # Extract headers from the first row (assuming the first row contains headers)
                headers = [th.text.strip() for th in rows[0].find_all("th")]
                headers.insert(0, "Property ID")  # Add Property ID as the first column

                # Extract data from the remaining rows
                table_data = []
                for row in rows[1:]:  # Skip the header row
                    # Extract Property ID from the Details link
                    details_link = row.find("a", href=True)
                    property_id = (
                        re.search(r"\d+$", details_link["href"]).group()
                        if details_link
                        else "N/A"
                    )

                    # Extract row data
                    row_data = [property_id] + [
                        td.text.strip() for td in row.find_all("td")
                    ]
                    table_data.append(row_data)

                # Create a DataFrame
                df = pd.DataFrame(table_data, columns=headers)

                # Write to Excel
                excel_filename = f"{countyName}_salesListing.xlsx"
                with pd.ExcelWriter(excel_filename, engine="xlsxwriter") as writer:
                    df.to_excel(writer, index=False, sheet_name="Sales Listings")

                    for column in df:
                        column_length = max(
                            df[column].astype(str).map(len).max(), len(column)
                        )
                        col_idx = df.columns.get_loc(column)
                        writer.sheets["Sales Listings"].set_column(
                            col_idx, col_idx, column_length
                        )

                print(f"Data exported to {excel_filename}")
            else:
                print("No table found.")
        else:
            print(f"Failed to fetch data. Status code: {response.status_code}")
    else:
        print("No county ID provided.")


def exportAllSalesListingData():
    xlsx_filename = f"Listings - {str(date.today())}.xlsx"

    # Create a new Excel file using xlsxwriter
    with xlsxwriter.Workbook(xlsx_filename) as workbook:
        # Fetch data for each county in parallel
        with ThreadPoolExecutor(
            max_workers=8
        ) as executor:  # Adjust max_workers based on your system's capabilities
            futures = {
                executor.submit(getSalesListingData, county["countyId"]): county["name"]
                for county in counties
            }

            for future in as_completed(futures):
                county_name = futures[future]
                try:
                    data = future.result()  # Get the result of the scraping task
                    if not data or not data.get("data"):
                        print(f"No data found for {county_name}")
                        continue

                    # Create a new sheet for the county
                    sheet_name = county_name[
                        :31
                    ]  # Sheet names are limited to 31 characters
                    worksheet = workbook.add_worksheet(sheet_name)

                    # Write the header row
                    headers = data.get("column_names", [])
                    for col_idx, header in enumerate(headers):
                        worksheet.write(0, col_idx, header)

                    # Write the data rows
                    for row_idx, row_data in enumerate(
                        data["data"], start=1
                    ):  # Start from row 1 (after header)
                        for col_idx, value in enumerate(row_data):
                            worksheet.write(row_idx, col_idx, value)

                    # Auto-adjust column widths
                    for col_idx, header in enumerate(headers):
                        max_length = max(
                            len(str(header)),  # Header length
                            max(
                                len(str(row_data[col_idx]))
                                for row_data in data[
                                    "data"
                                ]  # Longest value in the column
                            ),
                        )
                        worksheet.set_column(
                            col_idx, col_idx, max_length + 2
                        )  # Add padding

                    # Hyperlink the first column (Property ID)
                    link_format = workbook.add_format(
                        {"font_color": "blue", "underline": 1}
                    )
                    for row_idx, prop_id in enumerate(data["data"], start=1):
                        if (
                            prop_id and str(prop_id[0]).isdigit()
                        ):  # Ensure it's a valid ID
                            link = f"https://salesweb.civilview.com/Sales/SaleDetails?PropertyId={prop_id[0]}"
                            worksheet.write_url(
                                row_idx, 0, link, link_format, str(prop_id[0])
                            )

                except Exception as e:
                    print(f"Error processing {county_name}: {e}")

    print(f"Excel file '{xlsx_filename}' created successfully.")


def exportAllSalesListingDetailsDataFromCounty():
    xlsx_filename = f"Properties - {str(date.today())}.xlsx"

    with xlsxwriter.Workbook(xlsx_filename) as workbook:
        with ThreadPoolExecutor(
            max_workers=8
        ) as executor:  # Adjust max_workers based on your system's capabilities
            futures = {
                executor.submit(getPropertyDetailsByCounty, county["countyId"]): county[
                    "name"
                ]
                for county in small_counties
            }

            for future in as_completed(futures):
                county_name = futures[future]
                try:
                    data = future.result()  # Get the result of the scraping task
                    if not data:
                        print(f"No data found for {county_name}")
                        continue

                    worksheet = workbook.add_worksheet(
                        county_name[:31]
                    )  # Sheet name must be <= 31 chars

                    headers = list(
                        data[0].keys()
                    )  # Extract headers from the first row of data
                    for col_idx, header in enumerate(headers):
                        worksheet.write(0, col_idx, header)

                    # Write the data rows
                    for row_idx, row_data in enumerate(
                        data, start=1
                    ):  # Start from row 1 (after header)
                        for col_idx, key in enumerate(headers):
                            worksheet.write(row_idx, col_idx, row_data.get(key, ""))

                    for col_idx, header in enumerate(headers):
                        max_length = max(
                            len(str(header)),
                            max(
                                len(str(row_data.get(header, ""))) for row_data in data
                            ),
                        )
                        worksheet.set_column(col_idx, col_idx, max_length + 2)

                except Exception as e:
                    print(f"Error processing {county_name}: {e}")

    print(f"Scraping complete! Data saved to {xlsx_filename}")


if __name__ == "__main__":
    exportAllSalesListingData()
    exportAllSalesListingDetailsDataFromCounty()

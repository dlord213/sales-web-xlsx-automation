from data.counties_id import counties
import requests
from bs4 import BeautifulSoup
import re
import csv
import pandas as pd
from datetime import date


def extract_property_id(td):
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

        property_id = extract_property_id(
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


def getPropertyIdData(id: int):
    """
    Not yet finished.
    """
    url = f"https://salesweb.civilview.com/Sales/SaleDetails?PropertyId={id}"
    headers = {"User-Agent": "Mozilla/5.0"}
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        soup = BeautifulSoup(response.text, "html.parser")
        rows = soup.find("div", class_="table-responsive").find_all("tr")


def exportSalesListingData(id: int):
    """
    This function is to export the data of every sales listing in a county depending on the id at the args.
    """
    csv_headers = [
        "Property ID",
        "Sheriff #",
        "Sales Date",
        "Plaintiff",
        "Defendant",
        "Address",
    ]

    if id:
        url = f"https://salesweb.civilview.com/Sales/SalesSearch?countyId={id}"
        headers = {"User-Agent": "Mozilla/5.0"}
        response = requests.get(url, headers=headers)

        if response.status_code == 200:
            soup = BeautifulSoup(response.text, "html.parser")

            countyName: str = ""

            for county in counties:
                if county["countyId"] == id:
                    countyName = county["name"]

            last_updated: str = (
                soup.find("small").getText().replace("last updated: ", "")
            )
            rows = soup.find("form").find_all("tr")
            table_data = [
                [re.search(r"\d+$", row.find("td").find("a")["href"]).group()]
                + [td.text.strip() for td in row.find_all("td")[1:]]
                for row in rows
                if row.find("td") and row.find("td").find("a")
            ]

            with open(f"{countyName}_salesListing.csv", mode="w", newline="") as file:
                writer = csv.writer(file)
                writer.writerow(csv_headers)
                writer.writerows(table_data)


def exportAllSalesListingData():
    xlsx_filename = f"{str(date.today())}.xlsx"
    writer = pd.ExcelWriter(xlsx_filename, engine="xlsxwriter")
    county_data = {}

    # Fetch data for each county
    for county in counties:
        data = getSalesListingData(county["countyId"])

        if not data or not data.get("data"):
            print(f"No data found for {county['name']}")
            continue

        county_data[county["name"]] = data

    # Write data to Excel
    for county, data in county_data.items():
        if "column_names" not in data or not data["column_names"]:
            continue

        df = pd.DataFrame(data["data"], columns=data["column_names"])
        df = df.drop_duplicates()

        sheet_name = county[:31]  # Sheet names are limited to 31 characters

        df.to_excel(writer, sheet_name=sheet_name, index=False)

        worksheet = writer.sheets[sheet_name]
        workbook = writer.book
        link_format = workbook.add_format({"font_color": "blue", "underline": 1})

        # Set column width dynamically
        for idx, col in enumerate(df.columns):
            max_length = max(df[col].astype(str).map(len).max(), len(col))
            worksheet.set_column(idx, idx, max_length)

        # Always hyperlink the first column (Property ID)
        for row_num, prop_id in enumerate(df.iloc[:, 0], start=1):  # First column only
            if pd.notna(prop_id) and str(prop_id).isdigit():  # Ensure it's a valid ID
                link = f"https://salesweb.civilview.com/Sales/SaleDetails?PropertyId={prop_id}"
                worksheet.write_url(row_num, 0, link, link_format, str(prop_id))

    writer.close()
    print(f"Excel file '{xlsx_filename}' created successfully.")


if __name__ == "__main__":
    exportAllSalesListingData()

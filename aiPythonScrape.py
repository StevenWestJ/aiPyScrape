import os.path
import sys
import requests
from bs4 import BeautifulSoup
import pandas as pd
import xml.etree.ElementTree as ET
import time
from tqdm import tqdm
import re
from tkinter import filedialog, Tk
import logging
import colorlog

class Kirke:
    def __init__(self):
        self.kirke_id = None
        self.kirke_navn = None
        self.kirke_addr1 = None
        self.kirke_addr2 = None
        self.kirke_postnr = None
        self.kirke_by = None
        self.kirke_lat = None
        self.kirke_lng = None
        self.sogne_id = None
        self.sogne_navn = None
        self.sogndk_url= None
        self.provsti_id = None
        self.provsti_navn = None
        self.staff = []
        self.account_status = ""

'''
<div class="person_data">
    <div class="stilling pt-md-4 bigger-font"><font><font>Parish Priest (Church Bookkeeper)</font></font></div>
    <div class="navn"><font><font>Jesper Bacher</font></font></div>
    <div class="adr1"><font><font>Rubbeløkkevej 10</font></font></div>
    <div class="postnr_by"><font><font>4970 Rødby</font></font></div>
    <div class="email"><a><font><font>jeba@km.dk</font></font></a></div>
    <div class="tlf"><font><font>Phone: 54608118</font></font></div>
    <div class="my-6"><a><font><font>Secure inquiry</font></font></a></div>
</div>
'''
class Staff:
    def __init__(self):
        self.stilling = None
        self.navn = None
        self.adr1 = None
        self.postnr_by = None
        self.email = None
        self.tlf = None


def get_xml_data(url, logger):
    try:
        with requests.sessions.Session() as session:
            response = session.get(url)
            response.raise_for_status()
            return response.text
    except requests.exceptions.RequestException as e:
        logger.error("An error occurred while retrieving the data: %s", e)
        return None

def parse_kirke_xml(xml_data, kirker):
    root = ET.fromstring(xml_data)
    for kirke in root.findall("./kirke"):
        k = Kirke()
        k.kirke_id = int(kirke.find("kirkeId").text)
        k.kirke_navn = kirke.find("kirkenavn").text
        k.kirke_addr1 = kirke.find("kirkeaddr1").text
        k.kirke_addr2 = kirke.find("kirkeaddr2").text
        k.kirke_postnr = int(kirke.find("kirkepostnr").text)
        k.kirke_by = kirke.find("kirkeby").text
        k.kirke_lat = kirke.find("lat").text
        k.kirke_lng = kirke.find("lng").text
        k.provsti_id = int(kirke.find("provstiId").text)
        k.provsti_navn = kirke.find("provstinavn").text
        k.sogne_id = int(kirke.find("sogneId").text)
        k.sogne_navn = kirke.find("sognenavn").text
        k.sogndk_url = kirke.find("sogndkurl").text
        kirker.append(k)

def get_text_or_empty(element):
    return element.text if element else ""

def scrape_priests(kirke, logger):
    try:
        with requests.sessions.Session() as session:
            # Get the page content using requests
            page = session.get(kirke.sogndk_url + "praester-medarb")
            page.raise_for_status()
            soup = BeautifulSoup(page.content, "html.parser")
    except requests.exceptions.RequestException as e:
        logger.error("An error occurred while retrieving the web page for Kirke ID %s: %s", kirke.kirke_id, e)
        return

    try:
        # Find all the elements with class "person_data"
        staff_list = soup.find_all(class_="person_data")

        # Extract information about each staff member from the 'person_data' class
        kirke.staff = []
        for staff in staff_list:
            new_staff = Staff()
            new_staff.navn = get_text_or_empty(staff.find(class_="navn"))
            new_staff.stilling = get_text_or_empty(staff.find(class_="stilling"))
            new_staff.adr1 = get_text_or_empty(staff.find(class_="adr1"))
            new_staff.postnr_by = get_text_or_empty(staff.find(class_="postnr_by"))
            new_staff.email = get_text_or_empty(staff.find(class_="email"))
            new_staff.tlf = get_text_or_empty(staff.find(class_="tlf"))
            kirke.staff.append(new_staff)

    except requests.exceptions.RequestException as e:
        logger.error("An error occurred while scraping the staff data for Kirke ID %s: %s", kirke.kirke_id, e)

def save_to_excel(kirker, logger):
    # Check if user wants to save data
    save_file_choice = input("Do you want to save the data to an Excel file? (y/n) ")
    while save_file_choice not in ["y", "n"]:
        logger.warning("Invalid choice. Please try again.")
        save_file_choice = input("Do you want to save the data to an Excel file? (y/n) ")
    if save_file_choice == "n":
        logger.info("Data not saved.")
        return

    # Open file dialog to select file path
    root = Tk()
    root.withdraw()

    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
    if not file_path:
        logger.warning("No file selected. Data not saved.")
        return

    # Create a new Excel workbook
    workbook = pd.ExcelWriter(file_path, engine='openpyxl')

    # Create worksheet for Kirker and Staff data
    kirker_list = []
    for k in kirker:
        for s in k.staff:
            kirke_dict = {
                'Account Status': k.account_status,
                'Sogne_id': k.sogne_id,
                'Kirke_id': k.kirke_id,
                'Kirke_navn': k.kirke_navn
            }
            kirke_dict.update(k.__dict__)
            kirke_dict.update(s.__dict__)
            kirker_list.append(kirke_dict)

    kirker_df = pd.DataFrame(kirker_list)
    kirker_df.drop(['staff'], axis=1, inplace=True)
    kirker_df.to_excel(workbook, index=False, sheet_name='Kirker and Staff')

    # Save workbook to file path
    workbook.save()
    logger.info("Data saved to %s" % file_path)

def main(kirker):
    # Create logger and formatter
    handler = colorlog.StreamHandler()
    handler.setFormatter(colorlog.ColoredFormatter(
        "%(log_color)s%(levelname)s:%(message)s",
        log_colors={
            'DEBUG': 'cyan',
            'INFO': 'white',
            'WARNING': 'yellow',
            'ERROR': 'red',
            'CRITICAL': 'red,bg_white',
        }))
    logger = colorlog.getLogger()
    logger.setLevel(logging.DEBUG)
    logger.addHandler(handler)

    while True:
        logger.info("Press 1 to scrape new data from the web.")
        logger.info("Press 2 to import 'Account Status' field from an Excel file.")
        logger.info("Press E to exit.")
        choice = input("Enter your choice: ")

        # Get the arguments passed to the script
        args = sys.argv
        if choice == "1":
            xml_data = get_xml_data("http://sogn.dk/xmlfeeds/kirker.php", logger)
            if xml_data:
                parse_kirke_xml(xml_data, kirker)
                logger.info("%s churches found.", len(kirker))
                scrape_priests_choice = input("Do you want to scrape information about priests for each church? (Y/n) ")
                while scrape_priests_choice not in ["y", "n"]:
                    logger.warning("Invalid choice. Please try again.")
                    scrape_priests_choice = input("Do you want to scrape information about priests for each church? (Y/n) ")
                if scrape_priests_choice == "y":
                    for k in tqdm(kirker, total=len(kirker), desc="Scraping Priests Data"):
                        scrape_priests(k, logger)
                        time.sleep(0.5)
            else:
                logger.error("Unable to retrieve data from the web. Please try again.")

        elif choice == "2":
            if "--arg2" in args:
                logger.debug("Running in debug mode")
                # Get the index of the argument
                arg_index = args.index("--arg2")
                # Get the value of the argument
                arg_value = args[arg_index + 1]
                logger.debug(f"The value of --arg2 is: {arg_value}")
                file_path_3 = arg_value
                if not os.path.isfile(file_path_3):
                    logger.warning("File not found. Please try again or check this path exists: {}", arg_value)
                    continue
            else:
                # Code to run when not in debug mode
                root = Tk()
                root.withdraw()
                file_path_3 = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
                if not file_path_3:
                    logger.warning("No file selected. Please try again.")
                    continue
                if not os.path.isfile(file_path_3):
                    logger.warning("File not found. Please try again.")
                    continue

            # Read the Excel file
            try:
                df = pd.read_excel(file_path_3)
                logger.info("%s rows loaded from %s", len(df.index), file_path_3)
            except Exception as e:
                logger.error("Error reading Excel file: %s", str(e))

            # Update the account status of the Kirke objects based on the data in the DataFrame
            for index, row in df.iterrows():
                # 'CCLI Num' has sogne_id values.  This cell can have multiple sogne_ids
                ccli_num = row['CCLI Num']
                account_status = row['Account Status']
                if pd.notna(ccli_num):
                    ccli_nums = list(ccli_num.split(';'))

                    # Iterate over each ccli_num value
                    for num in ccli_nums:
                        # Define a lambda function to check if a Kirke object's sogne_id contains the ccli_num value
                        matching_function = lambda kirke: num in str(kirke.sogne_id)

                        # Use filter() to get a list of the matching Kirke objects
                        matching_kirker_num = list(filter(matching_function, kirker))

                        # Update the account status of the matching Kirke objects
                        for kirke in matching_kirker_num:
                            kirke.account_status = account_status

            save_to_excel(kirker, logger) 

        elif choice == "E":  
            save_to_excel(kirker, logger)
            sys.exit(0)

if __name__ == '__main__':
    kirker = []
    main(kirker)
import os.path
import sys
import requests
from bs4 import BeautifulSoup
from pandas.io.json import json_normalize
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
        self.sogne_id = None
        self.sogne_navn = None
        self.sogndk_url= None
        self.provsti_id = None
        self.provsti_navn = None
        self.priests = []
        self.account_status = ""

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
        k.provsti_id = int(kirke.find("provstiId").text)
        k.provsti_navn = kirke.find("provstinavn").text
        k.sogne_id = int(kirke.find("sogneId").text)
        k.sogne_navn = kirke.find("sognenavn").text
        k.sogndk_url = kirke.find("sogndkurl").text
        kirker.append(k)
   ## return kirker

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
        # Find all the elements with class "praester"
        priests = soup.find_all(class_="praester")

        # Extract information about each priest from the 'person' class
        kirke.priests = [{
            "name": get_text_or_empty(person_data.find(class_="navn")),
            "job": get_text_or_empty(person_data.find(class_="stilling")),
            "phone": "".join(re.findall(r'\d+', get_text_or_empty(person_data.find(class_="tlf")))),
            "email": get_text_or_empty(person_data.find(class_="email"))
        } for priest in priests
            if (person := priest.find(class_="person")) is not None
            and (person_data := person.find(class_="person_data")) is not None]

    except requests.exceptions.RequestException as e:
        logger.error("An error occurred while scraping the priests data for Kirke ID %s: %s", kirke.kirke_id, e)

def save_to_excel(kirker, logger):
    # Check if user wants to save data
    save_file_choice = input("Do you want to save the data to an Excel file? (Y/n) ")
    while save_file_choice not in ["Y", "n"]:
        logger.warning("Invalid choice. Please try again.")
        save_file_choice = input("Do you want to save the data to an Excel file? (Y/n) ")
    if save_file_choice == "n":
        logger.info("Data not saved.")
        return

    # Open file dialog to select file path
    root = Tk()
    root.withdraw()

    #if __debug__:
    #    file_path = "kirker_debug.xlsx"
    #else:
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
    if not file_path:
        logger.warning("No file selected. Data not saved.")
        return

    # Create a new Excel workbook
    workbook = pd.ExcelWriter(file_path, engine='openpyxl')

    # Create worksheet for Kirker data
    kirker_df = pd.DataFrame([k.__dict__ for k in kirker])
    kirker_df.drop('priests', axis=1, inplace=True)
    kirker_df.to_excel(workbook, index=False, sheet_name='Kirker')

    # Create worksheet for Priests data
    priests_df = pd.DataFrame()
    for k in kirker:
        for p in k.priests:
            p['kirke_id'] = k.kirke_id
            p['sogne_id'] = k.sogne_id
        priests_df = pd.concat([priests_df, pd.DataFrame(k.priests)])
    priests_df.to_excel(workbook, index=False, sheet_name='Priests')

    # Create a new list for Kirke objects without account status
    kirker_without_account_status = [k for k in kirker if k.account_status == ""]

    # Create worksheet for missing Account Status
    df_kirker_without_account_status = pd.DataFrame([k.__dict__ for k in kirker_without_account_status])
    df_kirker_without_account_status.drop('priests', axis=1, inplace=True)
    df_kirker_without_account_status.to_excel(workbook, index=False, sheet_name='No Account Status')

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
        logger.info("Press 1 to import existing data from an Excel file.")
        logger.info("Press 2 to scrape new data from the web.")
        logger.info("Press 3 to import 'Account Status' field from an Excel file.")
        logger.info("Press E to exit.")
        choice = input("Enter your choice: ")

        if choice == "1":
           # Get the arguments passed to the script
            args = sys.argv
            if "--arg1" in args:
                logger.debug("Running in debug mode")
                # Get the index of the argument
                arg_index = args.index("--arg1")
                # Get the value of the argument
                arg_value = args[arg_index + 1]
                logger.debug(f"The value of --arg1 is: {arg_value}")
                file_path_1 = arg_value
                if not os.path.isfile(file_path_1):
                    logger.warning("File not found. Please try again or check this path exists: {}", arg_value)
                    continue
            else:
                # Create a Tkinter window to use for the file dialog
                root = Tk()
                root.withdraw()

                # Open a file dialog to select the Excel file
                file_path_1 = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx;*.xls")])

                # Check that a file was selected
                if not file_path_1:
                    logger.warning("No file selected. Please try again.")
                    continue

                # Check that the file exists
                if not os.path.isfile(file_path_1):
                    logger.warning("File not found. Please try again.")
                    continue                
            try:
                df = pd.read_excel(file_path_1)
                logger.info("%s rows loaded from %s", len(df.index), file_path_1)

                for index, row in df.iterrows():
                    # Check if kirke_id already exists in kirker list
                    existing_kirke = next((k for k in kirker if k.kirke_id == row['kirke_id']), None)

                    if existing_kirke:
                        # Update existing kirke object
                        existing_kirke.kirke_id = row['kirke_id']
                        existing_kirke.kirke_navn = row['kirke_navn']
                        existing_kirke.kirke_addr1 = row['kirke_addr1']
                        existing_kirke.kirke_addr2 = row['kirke_addr2']
                        existing_kirke.kirke_postnr = row['kirke_postnr']
                        existing_kirke.kirke_by = row['kirke_by']
                        existing_kirke.sogne_navn = row['sogne_navn']
                        existing_kirke.sogndk_url = row['sogndk_url']
                        existing_kirke.provsti_id = row['provsti_id']
                        existing_kirke.provsti_navn = row['provsti_navn']
                        logger.info("Updated Kirke with kirke_id %s", row['kirke_id'])
                    else:
                        # Create new kirke object
                        kirke = Kirke()
                        kirke.kirke_id = row['kirke_id']
                        kirke.kirke_navn = row['kirke_navn']
                        kirke.kirke_addr1 = row['kirke_addr1']
                        kirke.kirke_addr2 = row['kirke_addr2']
                        kirke.kirke_postnr = row['kirke_postnr']
                        kirke.kirke_by = row['kirke_by']
                        kirke.sogne_id = row['sogne_id']
                        kirke.sogne_navn = row['sogne_navn']
                        kirke.sogndk_url = row['sogndk_url']
                        kirke.provsti_id = row['provsti_id']
                        kirke.provsti_navn = row['provsti_navn']
                        kirker.append(kirke)
                        logger.info("Added new Kirke with kirke_id %s", row['kirke_id'])

            except FileNotFoundError:
                logger.error("File not found. Please try again.")
            except Exception as e:
                logger.error("An error occurred while loading the data: %s. Please try again.", e)


        elif choice == "2":
            xml_data = get_xml_data("http://sogn.dk/xmlfeeds/kirker.php", logger)
            if xml_data:
                parse_kirke_xml(xml_data, kirker)
                logger.info("%s churches found.", len(kirker))
                scrape_priests_choice = input("Do you want to scrape information about priests for each church? (Y/n) ")
                while scrape_priests_choice not in ["Y", "n"]:
                    logger.warning("Invalid choice. Please try again.")
                    scrape_priests_choice = input("Do you want to scrape information about priests for each church? (Y/n) ")
                if scrape_priests_choice == "Y":
                    for k in tqdm(kirker, total=len(kirker), desc="Scraping Priests Data"):
                        scrape_priests(k, logger)
                        time.sleep(0.5)

                        # Backup kirker list to Excel file
                    df = pd.DataFrame([k.__dict__ for k in kirker])
                    df.to_excel("kirker_backup.xlsx", index=False, sheet_name='Kirker', engine='openpyxl', startrow=0, header=True)
                    logger.info("Kirker list backed up to 'kirker_backup.xlsx'.")
            else:
                logger.error("Unable to retrieve data from the web. Please try again.")

        elif choice == "3":
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
from PyPDF2 import PdfReader
from PyPDF2._page import PageObject
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
import gspread
import re
import os
import json
import datetime
import warnings
import glob
import sys
import logging

class Settings:
	def __init__(self, pdfs_folder_path: str = "", google_sheet_url: str = ""):
		self.pdfs_folder_path = pdfs_folder_path
		self.google_sheet_url = google_sheet_url

class Decklist:
	def __init__(self, player: str, deck:str, cards: str):
		self.player = player
		self.deck = deck
		self.cards = cards

# Global Variables
#region
google_sheet_path_re_pattern = r"https://docs.google.com/spreadsheets/d/([a-zA-Z0-9-_]+)"
gspread_credential_path = "credential.json"
settings_file_path = "settings.json"
settings = Settings()
apostrophe_in_html = "&#39;"
decklist_correct_first_line = "DECK REGISTRATION SHEETTable"
decklist_unwanted_lines = { "Main Deck Continued:", "# in deck: Card Name:"}
decklists: list[Decklist] = []
decklists_discarded: list[str] = []
driver: webdriver = None
#endregion

# Common Methods
#region
def ask_question_with_input_validation(question:str, validInputs:set[set]):
	print (question)
	while True:
		answer = input ()
		if (answer.strip().lower() in validInputs):
			return answer.strip().lower()
		else:
			print("Invalid input. Please enter one of these:", validInputs)

def get_full_path(relative_path):
    """ Get the absolute path to a resource, works for PyInstaller's onefile mode """
    if hasattr(sys, '_MEIPASS'):
        # If running as a PyInstaller bundle, locate in the temp directory
        return os.path.join(sys._MEIPASS, relative_path)
    else:
        # When running as a script, locate relative to the script
        return os.path.join(os.path.abspath("."), relative_path)
#endregion

# PDF Reading
#region
def replace_html_entities (text:str):
	return text.replace(apostrophe_in_html, "'")

def replace_unwanted_lines (text:str):
	for unwantedLine in decklist_unwanted_lines:
		text = text.replace(unwantedLine, "")
	return text

def find_text_between_two (text:str, start:str, end:str, dotall_pattern:bool = False):
	pattern = f'(?<={re.escape(start)})(.*?)(?={re.escape(end)})'

	if (dotall_pattern):
		match = re.search(pattern, text, re.DOTALL)
	else:
		match = re.search(pattern, text)

	if (match):
		return match.group(1).strip()
	else:
		return None

def add_space_after_number_of_line (line:str):
	pattern = r'^(\d+)(.*)'
	match = re.match(pattern, line)

	if (match):
		return match.group(1) + ' ' + match.group(2)
	else:
		return None

def add_space_after_number_of_all_lines (text:str):
	lines = text.splitlines()
	for i in range(len(lines)):
		lines[i] = add_space_after_number_of_line(lines[i])

	return '\n'.join(lines)

def trim_text_after_character (text:str, char:str):
	lines = text.splitlines()
	for i in range(len(lines)):
		lines[i] = lines[i].split(char)[0].strip()

	return '\n'.join(lines)

def clean_cards_list(text:str):
	text = add_space_after_number_of_all_lines(text)
	text = trim_text_after_character(text, '(')
	return text

def get_name (text:str):

	start = "First Name:"
	end = "DCI #"
	return find_text_between_two(text, start, end)

def get_last_name (text:str):
	start = "Last Name:"
	end = "First Name:"
	return find_text_between_two(text, start, end)

def get_deck_name (text:str):
	start = "Deck Name:"
	end = "\n"
	return find_text_between_two(text, start, end)

def get_main_cards (text:str):
	start = "Main Deck:"
	end = "Sideboard:"
	text = find_text_between_two(text, start, end, dotall_pattern = True)
	return clean_cards_list(text)

def get_sideboard_cards (text:str):
	start = "Sideboard:"
	end = "Total Number of Cards in Main Deck:"
	text = find_text_between_two(text, start, end, dotall_pattern = True)
	return clean_cards_list(text)

def get_pdfs_folder_path_from_input():
	settings.pdfs_folder_path = input("Enter PDF folder path: ")

def get_pdfs_folder_path_from_settings():
	print(settings.pdfs_folder_path + "\nDo you want to use this path? (Empty = Yes, Else = No)")
	if (input() != ""):
		get_pdfs_folder_path_from_input()

def get_pdf_files():
	print("Starting to read PDFs...")				
	while True:
		pathIsEmpty = settings.pdfs_folder_path == None or settings.pdfs_folder_path == ""
		if (pathIsEmpty):
			get_pdfs_folder_path_from_input()
		else:
			get_pdfs_folder_path_from_settings()
			
		files = []
		for root, dirs, filenames in os.walk(settings.pdfs_folder_path):
			for filename in filenames:
				if filename.endswith('.pdf'):
					files.append(os.path.join(root, filename))

		if (len(files) == 0):
			print("No PDF files were found. Please try again.")
			continue
		break
			
	write_settings_file()
	return files

def has_valid_decklist(pdf: str):
	reader = PdfReader(pdf)
	if (len(reader.pages) == 0):
		decklists_discarded.append(pdf + " The file is empty.")
		return False

	text = reader.pages[0].extract_text()
	if (text is None or len(text) == 0 or text.startswith(decklist_correct_first_line) == False):
		decklists_discarded.append(pdf + " The file is in an incorrect format.")
		return False

	return True

def create_decklist_from_text (text: str):
	text = replace_html_entities(text)
	text = replace_unwanted_lines(text)
	name = get_name(text)
	name += " "
	name += get_last_name(text)
	deck = get_deck_name(text)
	cards = get_main_cards(text)
	cards += ("\nSIDEBOARD\n")
	cards += get_sideboard_cards(text)
	decklist = Decklist(name, deck, cards)
	decklists.append(decklist)
#endregion

# Settings
#region
def write_settings_file():
	with open(settings_file_path, 'w') as file:
		json.dump(vars(settings), file, indent=4)
#endregion

# Google Sheet
#region
def get_google_sheet_url_from_input():
	while True:
		settings.google_sheet_url = input("Enter Google Sheet url: ")
		isNotAGoogleSheetUrl = re.match(google_sheet_path_re_pattern, settings.google_sheet_url) is None
		
		if (isNotAGoogleSheetUrl):
			print("It is not a Google Sheet URL. Please try again.")
			continue
		else:
			break

	write_settings_file()

def get_google_sheet_url_from_settings():
	print(settings.google_sheet_url + "\nDo you want to use this url? (Empty = Yes, Else = No)")
	if (input() != ""):
		get_google_sheet_url_from_input()

def get_google_sheet_url():
	urlIsEmpty = settings.google_sheet_url == None or settings.google_sheet_url == ""

	if (urlIsEmpty):
		get_google_sheet_url_from_input()
	else:
		get_google_sheet_url_from_settings()

def get_unique_worksheet_name(names:list[str]):
	now = datetime.datetime.now()
	date = f"{now.day}/{now.month}/{now.year}"
	name = date
	index = 2

	while name in names:
		name = f"{date} ({index})"
		index += 1

	return name

def try_filling_google_sheet():
	try:
		print("Getting credentials...")
		path = get_full_path(gspread_credential_path)
		with open(path, 'r') as file:
			credential = json.load(file)

		print("Accessing the Google Sheet...")
		gc = gspread.service_account_from_dict(credential)
		spreadsheet = gc.open_by_url(settings.google_sheet_url)
		
		print("Filling the Google Sheet...")
		worksheetNames = existing_worksheets = [worksheet.title for worksheet in spreadsheet.worksheets()]
		worksheetName = get_unique_worksheet_name(worksheetNames)		
		index = len(spreadsheet.worksheets())
		worksheet = spreadsheet.add_worksheet(worksheetName, rows=1000, cols=200, index=index)
		col = 1
		index = 1
		total = len(decklists)
		for decklist in decklists:
			if (index < total):
				print(f"Processing decklists {index}/{total}", end='\r')
			else:
				print(f"Processing decklists {index}/{total}")
			worksheet.update_cell(1, col, decklist.player)
			worksheet.update_cell(2, col, decklist.cards)
			col += 1
			index += 1

		print ("Formatting the Google Sheet...")
		worksheet.format("A:Z", {"verticalAlignment": "TOP"})
		print ("DONE!\n")
	except Exception as e:
		print (f"An unexpected error ocurred.\n{e}.")
#endregion

# Web Navigation
#region
def get_driver():
	try:
		print("Opening Chrome...")
		global driver
		driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
		driver.minimize_window()
	except:
		print("Google Chrome is missing in this machine. This script only supports Chrome.")

def try_connmection_callback(driver: webdriver):
	return driver.execute_script('return document.readyState') == 'complete'		

def try_connection(url:str):
	try:
		print("Connecting with:", url)
		driver.get(url)
		WebDriverWait(driver, 10).until(try_connmection_callback)
	except TimeoutException:
		print("Timed out waiting for page to load. Try again later.")
	except Exception as e:
		print(f"An unexpected error occurred: {e}")
	print("Connected!")
#endregion

# MAIN
#region
def setup():
	warnings.filterwarnings("ignore", category=UserWarning, module="PyPDF2")	
	logging.getLogger("PyPDF2").setLevel(logging.CRITICAL)

def read_settings():
	if os.path.exists(settings_file_path):
		with open(settings_file_path) as file:
			data = json.load(file)
			structureIsCompatible = set(data.keys()) == set(vars(settings).keys()) #It won't be compatible if Settings get's updated and the JSON reflects an old structure
			if(structureIsCompatible):
				vars(settings).update(data)

def read_decklists():
	files = get_pdf_files()
	print("Reading the PDFs...")
	for file in files:
		if not (has_valid_decklist(file)):
			continue
		text = PdfReader(file).pages[0].extract_text()	
		extract_decklist(text)
	print("DONE!\n")

def fill_google_sheet():
	print("Preparing to access the Google Sheet...")
	try:
		get_google_sheet_url()
		try_filling_google_sheet()
	except:
		input("Asd")
		exit()
	finally:
		if (driver):
			driver.quit()

def log():
	if (len(decklists_discarded) > 0):
		print("The following files were not uploaded:")
		print("\n".join(decklists_discarded))
		input()
#endregion

setup()
read_settings()
read_decklists()
fill_google_sheet()
log()
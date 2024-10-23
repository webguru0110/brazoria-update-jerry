import time
import threading
import tkinter as tk
from multiprocessing import freeze_support
from tkinter import filedialog as fd
from tkinter import ttk
from urllib.parse import urlparse
import openpyxl
import requests
import sv_ttk
from bs4 import BeautifulSoup

# Open the Excel file
workbook = openpyxl.load_workbook(
    'Brazoria County Delinquent Tax Roll 2024 0510.xlsx')
sheet = workbook.active

# Select the Account number row for processing
account_numbers = [cell.value for cell in sheet['A'] if cell.value is not None]


def scrape_links(account_number, row_num):
    global info_text
    if not (row_num % 2000):
        workbook.save('update.xlsx')

    # formatted_number = str(account_number).zfill(11)
    url = f"https://tax.brazoriacountytx.gov/Accounts/AccountDetails?taxAccountNumber={account_number}"
    payload = {}
    response = requests.request("GET", url, data=payload, verify=True).text
    soup = BeautifulSoup(response, "html.parser")

    property_type_tag = soup.find_all(
        'div', class_='d-flex d-md-block justify-content-end')[1]
    property_type = property_type_tag.text.strip().lower()

    owner_name_address = soup.find_all('div', class_='mb-2 mb-md-0 col-6 col-md-3')[
        0].text.strip().lower()
    # property_street = soup.find_all('div', class_='mb-2 mb-md-0 col-md-3')[
    #     0].text.replace('Location', '').strip()
    # sheet[f'J{row_num}'].value = property_street.lstrip('0')
    # sheet[f'L{row_num}'].value = "TX"
    number_of_years = len(soup.find_all('div', class_='card'))-2
    sheet[f'AH{row_num}'].value = f"{number_of_years}"
    # Check if legal description refers to real estate or personal property
    if any(keyword in property_type for keyword in ['vehicles', 'personal', 'ff&e', 'mobile home', 'mineral']):
        sheet[f'AB{row_num}'].value = "Personal Property"
        return

    legal_description = soup.find_all(
        'div', class_='col-md-3')[2].find_all('div')[1].text.strip().lower()
    if any(keyword in legal_description for keyword in ['vehicles', 'personal property', 'ff&e', 'mobile', 'mineral']):
        sheet[f'C{row_num}'].value = "Personal Property"
        return
    # Check if owner's name contains certain keywords
    # if 'real' in property_type:
    #     sheet[f'N{row_num}'].value = "Real Estate"
    #     final_step(soup, row_num)

    if any(keyword in owner_name_address for keyword in ['estate', ' est ', 'deceased', 'unknown']):
        status = 'Estate Owned' if 'estate owned' in owner_name_address or 'est' in owner_name_address else 'Unknown' if 'unknown' in owner_name_address else ''
        sheet[f'B{row_num}'].value = status
        final_step(soup, row_num)
    else:

        # Find the <b> tag containing "Improvement Value:"
        improvement_value_tag = soup.find('div', class_='d-flex flex-column d-md-block').find(lambda tag: tag.name ==
                                                                                              'div' and 'Improvement:' in tag.text)
        if improvement_value_tag == None:
            improvement_value = 0
        else:
            # Get the next sibling, which contains the value
            improvement_value = float(
                improvement_value_tag.text.replace('Improvement:', '').replace('$', '').replace(',', '').strip())
        if improvement_value > 25000:
            status = 'Improvements'
            sheet[f'C{row_num}'].value = status
            # owner = soup.find_all('div', class_='mb-2 mb-md-0 col-6 col-md-3')[
            #     0]
            # owner_name = owner.find('span').text.strip()
            # sheet[f'D{row_num}'].value = owner_name
            # address = owner.text.replace(
            #     'Owner', '').replace(owner_name, '').strip()
            # lines = address.split('\n')

            # # Extract address, city, state, and zip code
            # address_line = lines[0]
            # city_state_zip = lines[1]

            # # Split city, state, and zip code
            # city, state_zip = city_state_zip.split(', ')
            # state, zip_code = state_zip.split(' ')

            # sheet[f'E{row_num}'].value = address_line
            # sheet[f'G{row_num}'].value = city
            # sheet[f'H{row_num}'].value = state
            # sheet[f'I{row_num}'].value = zip_code
            return
        else:
            final_step(soup, row_num)


def final_step(soup, row_num):
    # owner = soup.find_all('div', class_='mb-2 mb-md-0 col-6 col-md-3')[
    #     0]
    # owner_name = owner.find('span').text.strip()
    # sheet[f'D{row_num}'].value = owner_name
    # address = owner.text.replace('Owner', '').replace(owner_name, '').strip()
    # lines = address.split('\n')

    # # Extract address, city, state, and zip code
    # address_line = lines[0]
    # city_state_zip = lines[1]

    # # Split city, state, and zip code
    # city, state_zip = city_state_zip.split(', ')
    # state, zip_code = state_zip.split(' ')

    # sheet[f'E{row_num}'].value = address_line
    # sheet[f'G{row_num}'].value = city
    # sheet[f'H{row_num}'].value = state
    # sheet[f'I{row_num}'].value = zip_code
    # Find the tag containing "Land Value:"
    land_value_tag = soup.find('div', class_='d-flex flex-column d-md-block').find(lambda tag: tag.name ==
                                                                                   'div' and 'Land:' in tag.text)
    if land_value_tag == None:
        land_value = 0
    else:
        # Get the next sibling, which contains the value
        land_value = float(land_value_tag.text.replace(
            'Land:', '').replace('$', '').replace(',', '').strip())

    # Check land value
    if land_value < 15000:
        status = 'Low Value'
        sheet[f'C{row_num}'].value = status
        sheet[f'AC{row_num}'].value = f"{land_value :,}"
        return

    sheet[f'AC{row_num}'].value = f"{land_value :,}"
    # Find the <b> tag containing "Improvement Value:"
    improvement_value_tag = soup.find('div', class_='d-flex flex-column d-md-block').find(lambda tag: tag.name ==
                                                                                          'div' and 'Improvement:' in tag.text)
    if improvement_value_tag == None:
        improvement_value = 0
    else:
        # Get the next sibling, which contains the value
        improvement_value = float(
            improvement_value_tag.text.replace('Improvement:', '').replace('$', '').replace(',', '').strip())
    sheet[f'AD{row_num}'].value = f"{improvement_value :,}"

    # Get the next sibling, which contains the value
    taxes_due_tag = soup.find(
        'h5', class_='font-weight-light')
    # Get the next sibling, which contains the value
    taxes_due = float(taxes_due_tag.text.strip().replace(
        '$', '').replace(',', ''))
    sheet[f'AF{row_num}'].value = f"{taxes_due :,}"
    total_value = land_value + improvement_value
    sheet[f'AE{row_num}'].value = f"{total_value :,}"
    equity = total_value - taxes_due
    sheet[f'AG{row_num}'].value = f"{equity :,}"
    if equity < 25000:
        status = 'Low Equity'
        sheet[f'C{row_num}'].value = status
        return

    if sheet[f'B{row_num}'].value:
        property_type = "Land, Commercial, Residental"
    else:
        property_type = "Land"

    sheet[f'AB{row_num}'].value = property_type
    sheet[f'C{row_num}'].value = "Research"

    return


def process_account_numbers():
    threads = []
    max_threads = 100  # Maximum number of threads
    while len(threads) < max_threads:
        for idx, account_number in enumerate(account_numbers, 1):
            if idx == 1:
                continue
            print(idx, account_number)
            thread = threading.Thread(
                target=scrape_links, args=(account_number, idx))
            threads.append(thread)
            thread.start()
            if idx == len(account_numbers):
                for thread in threads:
                    thread.join()
                threads.clear()
                return

            # Adjust the sleep time as needed (e.g., 1 second)
            time.sleep(0.05)

    return


def main():
    global startbot
    global info_text

    startbot.config(state="disabled", text="Started...")

    process_account_numbers()
    # Save the modified workbook
    print("Processing Completed!")

    try:
        workbook.save(
            'Brazoria County Delinquent Tax Roll 2024 0510_update.xlsx')
        # Close the workbook
        workbook.close()
    except FileNotFoundError:
        info_text.config(text="file cannot found!")
        return

    startbot.config(state="enabled")
    info_text.config(text="Completed!")


if __name__ == '__main__':
    freeze_support()

    app = tk.Tk()
    app.title(f'Tarrant Tax Update')
    app.geometry('400x400')
    app.minsize(400, 400)
    app.maxsize(400, 400)

    ttk.Frame(app, height=30).pack()
    title = tk.Label(app, text='Tarrant Tax Update',
                     font=("Calibri", 24, "bold"))
    title.pack(pady=20)

    def select_file():
        global keywords_filepath

        file_path = fd.askopenfilename()
        file_path_short = file_path[file_path.rindex('/') + 1:]

        keywords_element.config(text=file_path_short)
        keywords_filepath = file_path

    keywords_filepath = ''

    keywords_info = ttk.Labelframe(app, text='Source File')
    keywords_info.pack(padx=60, pady=20)
    keywords_element = ttk.Button(keywords_info, text='Select file', width=40,
                                  command=lambda: select_file())
    keywords_element.pack(padx=10, pady=10, fill=tk.X)

    startbot = ttk.Button(app, text='Start Bot', style='Accent.TButton', width=15,
                          command=lambda: threading.Thread(target=main).start())
    startbot.pack(pady=10)

    info_text = ttk.Label(app, text='', justify=tk.CENTER)
    info_text.pack(pady=5)

    sv_ttk.set_theme('dark')
    app.mainloop()

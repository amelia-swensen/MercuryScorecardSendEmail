from send_email import send_outlook_email
import os
import pandas as pd
from bs4 import BeautifulSoup
import re
import datetime
import numpy as np


email_mappings = {'PC 128': '<amelia.swensen@tigris-fp.com>',
                  'PC 206': 'amelia.swensen@tigris-fp.com',
                  'PC 227': 'amelia.swensen@tigris-fp.com',
                  'PC 353': 'amelia.swensen@tigris-fp.com',
                  'PC 384': 'amelia.swensen@tigris-fp.com',
                  'PC 524': 'amelia.swensen@tigris-fp.com',
                  'PC 527': 'amelia.swensen@tigris-fp.com',
                  'PC 528': 'amelia.swensen@tigris-fp.com',
                  'PC 567': 'amelia.swensen@tigris-fp.com',
                  'PC 580': 'amelia.swensen@tigris-fp.com',
                  'PC 597': 'amelia.swensen@tigris-fp.com',
                  'PC 701': 'amelia.swensen@tigris-fp.com',
                  'PC 738': 'amelia.swensen@tigris-fp.com',
                  'PC 719': 'amelia.swensen@tigris-fp.com',
                  'PC 752': 'amelia.swensen@tigris-fp.com',
                  'PC 784': 'amelia.swensen@tigris-fp.com',
                  'PC 79': 'amelia.swensen@tigris-fp.com',
                  'PC 328': 'amelia.swensen@tigris-fp.com'}
# 'PC 79' : 'amelia.swensen@tigris-fp.com'
'''
email_mappings = {'PC 128': 'Wright, Bryan <bryan.wright@hajoca.com>; Mcguire, Eric <emcguire@hajoca.com>',
                  'PC 206': '<michaelw@bestplg.com>; <jonas@bestplg.com>Jonas Weiner; <phyllisl@bestplg.com>Phyllis Lupi; <johnr@bestplg.com>John Romagno Iii',
                  'PC 227': '<doug.caux@hajoca.com>Doug Caux; <valerie.puig@hajoca.com>Valerie Puig; <richard.yockel@hajoca.com>Richard Yockel',
                  'PC 353': '<bill.halliburton@abledistributing.com>Bill Halliburton; <ben.bergersen@hajoca.com>Ben Bergersen; <josh.kerslake@abledistributing.com>Josh Kerslake; <chris.haglund@hajoca.com>Christopher Haglund',
                  'PC 384': '<s.younggreen@martzsupply.com>; <b.scarpetta@martzsupply.com>Brian Scarpetta',
                  'PC 524': '<richard.mccandless@hughessupply.com>Richard McCandless; <lisa.misakian@hughessupply.com>Lisa Misakian; <mike.riccio@hughessupply.com>Mike Riccio',
                  'PC 527': '<eric.spraker@hajoca.com>Eric Spraker; <trevor.russell@hajoca.com>Trevor Russell; <matt.moore@hajoca.com>Matt Moore; <Paul.Ganger@hajoca.com>Paul Ganger',
                  'PC 528': '<Jackson.allen@hajoca.com>Jackson Allen; <Jessica.rogers@hughessupply.com>Jessica Rogers',
                  'PC 567': '<matthew.booth@hughessupply.com>Matthew Booth; <jeremy.pope@hughessupply.com>Jeremy Pope',
                  'PC 580': '<justin.kary@hajoca.com>Justin Kary; <jessica.godinez@hajoca.com>Jessica Godinez',
                  'PC 597': '<nate.devlin@hughessupply.com>Nate Devlin; <ryan.steinbeck@hajoca.com>Ryan Steinbeck; <nathaniel.columna@hajoca.com>Nathaniel Columna',
                  'PC 701': '<jj.donovan@mooresupply.com>JJ Donovan; <james.bell@mooresupply.com>James Bell',
                  'PC 738': '<Noah.Hardin@mooresupply.com>Noah Hardin; <tyler.oquin@mooresupply.com>Tyler O'Quin',
                  'PC 719': '<cbyrd@mooresupply.com>Chris Byrd; <Troy.Conyers@mooresupply.com>Troy Conyers; <Mike.Cottrell@facetshome.com>Mike Cottrell; <david.asbra@mooresupply.com>David Asbra; <rthiem@mooresupply.com>Randy Thiem',
                  'PC 752': '<Liz.williams@mooresupply.com>Elizabeth Williams; <Nathan.miller@mooresupply.com>Nathan Miller; <Tyler.hart@mooresupply.com>Tyler Hart; <Jose.elias@mooresupply.com>Jose Elias; <djwalthall@mooresupply.com>',
                  'PC 784': '<ctportier@tpwmail.com>Casey Portier; <kbelaire@tpwmail.com>Ken Belaire'}
'''


def get_files():
    current_date = datetime.datetime.now()
    iso_year, iso_week_number, iso_weekday = current_date.isocalendar()

    if iso_week_number == 1:
        year = iso_year - 1
        week = 52
    else:
        year = iso_year
        week = iso_week_number - 1

    folder_path = rf'N:\Project Mercury\Weekly PC Scorecards\{year}\W{week}'
    # largest_folder = get_largest_number_folder(parent_directory)

    '''if largest_folder:
        print("Folder with the largest number:", largest_folder)

        # Get the path to the folder with the largest number
        folder_path = os.path.join(parent_directory, rf'{largest_folder}')

        # Iterate through the files in the folder'''
    for file_name in os.listdir(folder_path):
        if 'INTERNAL' not in file_name and 'Project Mercury' not in file_name and '~$' not in file_name:
            end = file_name.find('Mercury') - 1
            pc_num = file_name[:end]
            email_addy = email_mappings[pc_num]
    for file_name in os.listdir(folder_path):
        if 'INTERNAL' not in file_name and 'Project Mercury' not in file_name and '~$' not in file_name:
            print(file_name)
            file_path = os.path.join(folder_path, file_name)
            end = file_name.find('Mercury')-1
            pc_num = file_name[:end]
            email_addy = email_mappings[pc_num]
            html = get_last_4(file_path)
            send_outlook_email(file_name, email_addy, file_path, html)
        else:
            print(f"Invalid file: {file_name}")


def get_largest_number_folder(parent_dir):
    largest_number = -1
    largest_folder = None

    # Iterate through all the folders in the parent directory
    for folder_name in os.listdir(parent_dir):
        if not os.path.isdir(os.path.join(parent_dir, folder_name)):
            continue

        # Extract the number from the folder name
        try:
            number = int(folder_name[1:])
        except ValueError:
            continue

        # Keep track of the largest number found
        if number > largest_number:
            largest_number = number
            largest_folder = folder_name
    print(largest_folder)
    return largest_folder


def get_last_4(file_path):
    data = pd.read_excel(file_path, skiprows=2, engine='openpyxl')

    if len(data.columns) >= 5:
        # Get the first column and the last four columns
        columns_to_include = [data.columns[0]] + data.columns[-4:].tolist()
    else:
        # Include all columns if the DataFrame has less than 5 columns
        columns_to_include = data.columns.tolist()

    selected_data = data[columns_to_include]
    for col in selected_data.columns[1:]:
        selected_data[col] = selected_data[col].astype('float').round(decimals=2)

    # Format money rows
    selected_data.loc[4, columns_to_include[1:]] = selected_data.loc[4, columns_to_include[1:]].apply('$ {:,.0f}'.format)
    selected_data.loc[5, columns_to_include[1:]] = selected_data.loc[5, columns_to_include[1:]].apply('$ {:,.0f}'.format)

    # Format Percent Rows
    selected_data.loc[6, columns_to_include[1:]] = selected_data.loc[6, columns_to_include[1:]].apply('{:.0%}'.format)

    selected_data.loc[0, columns_to_include[1:]] = selected_data.loc[0, columns_to_include[1:]].astype(int)
    selected_data.loc[1, columns_to_include[1:]] = selected_data.loc[1, columns_to_include[1:]].astype(int)
    selected_data.loc[2, columns_to_include[1:]] = selected_data.loc[2, columns_to_include[1:]].astype(int)
    selected_data.loc[3, columns_to_include[1:]] = selected_data.loc[3, columns_to_include[1:]].astype(int)
    selected_data.loc[7, columns_to_include[1:]] = selected_data.loc[7, columns_to_include[1:]].astype(int)
    selected_data.loc[8, columns_to_include[1:]] = selected_data.loc[8, columns_to_include[1:]].astype(int)
    selected_data.loc[9, columns_to_include[1:]] = selected_data.loc[9, columns_to_include[1:]].astype(int)

    html_table = selected_data.to_html(index=False)
    # print(selected_data)

    format = '''<head>
                <style>
                table {
                  border-collapse: collapse;
                }
                
                th {
                  padding: 5px;
                  text-align: center;
                }
                td {
                  padding-left: 5px;
                  padding-right: 5px;
                  text-align: right;
                }
                </style>
                </head>'''
    html_table = format + html_table
    soup = BeautifulSoup(html_table, "html.parser")

    words_to_match = [
        "Orders",
        "Lines",
        "Pieces",
        "Unique SKU Count",
        "PC Sales",
        "Average Ticket",
        "Service Failures \(% of orders\)",
        "Can't Find",
        "Damage",
        "Late Shipment",
    ]
    pattern = re.compile("|".join(words_to_match))
    indent_pattern = re.compile("|".join(words_to_match[-3:]))

    for tag in soup.find_all("td"):
        inner_text = tag.get_text().strip()
        if re.match(pattern, inner_text):
            tag['style'] = 'text-align: left;'
            if re.match(indent_pattern, inner_text):
                tag['style'] = 'text-align: left; padding-left: 25px;'

    html_table = soup.prettify()
    # print(html_table)

    return html_table
import os
import re
import sys
import glob
import time
import json
import fitz
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from sostocked_templates import sostocked_shipment
from xlsxwriter.utility import xl_rowcol_to_cell

"""
To do:
1. convert to ST packlist
2. fix SKU on Amazon box labels x 
3. create FBA shipment control sheet (PDF) to be sent to ST
"""

# determine if application is a script file or frozen exe
if getattr(sys, 'frozen', False):
    currentDirectory = os.path.dirname(os.path.realpath(sys.executable))
    print(f"Current Direcotry: {currentDirectory}")
elif __file__:
    currentDirectory = os.path.dirname(__file__)

# Folder locations
downloads_folder = os.path.join(os.environ.get('HOMEPATH'), 'Downloads')
amazonShipmentsDirectory = os.path.join(currentDirectory, 'Amazon Shipments')

# Check for dump folders
os.makedirs(amazonShipmentsDirectory, exist_ok=True)

# Master Data File
masterData = pd.read_excel('Master Data File.xlsx', sheet_name='All Products').dropna(how='all').fillna('')

def add_amazon_sku(doc, save_location):
    """
    Fixes Amazon's packing list with long SKUs that results to DIP226 - S...one 7 Pack
    Input: doc
    Output: PDF to designated folder
    """
    print("Fixing Amazon SKU")
    for page in doc:
        pageText = json.loads(page.get_text('json'))
        pageBlocks = pageText['blocks']

        # Getting FBA Box ID
        for block in pageBlocks:    # iterates through blocks, lines & spans
            try:
                for line in block['lines']:

                    for span in line['spans']:
                        if re.match(r'^FBA\w', span['text']):
                            FBAId = span['text']    # FBA Box ID
                            fbaBoxIdBlockNumber = block['number']
                            FBAIdbbox   = span['bbox']
                            # font properties of "Single SKU" block
                            SingleSKUnumber = fbaBoxIdBlockNumber + 1
                            SingleSKUspan   = pageBlocks[SingleSKUnumber]['lines'][0]['spans'][0]
                            # Fixes SKU
                            SKUnumber = fbaBoxIdBlockNumber + 2
                            SKUspan   = pageBlocks[SKUnumber]['lines'][0]['spans'][0]
                            SKU = SKUspan['text']
                            if '...' in SKU:
                                splittedSKU = SKU.split('...')[0]
                                SKU = masterData[masterData['SKU'].str.startswith(splittedSKU)]['SKU'].values[0]
    
                            # Writes SKU to left corner
                            print(f"\tWRITING SKU: {SKU}")
                            rotation = 90 if line['dir'] == [0, -1] else 0
                            textLength = fitz.get_text_length(SKU, span['font'], 8)
                            if rotation:
                                SKUbbox = [SingleSKUspan['bbox'][0], FBAIdbbox[1]-50, SingleSKUspan['bbox'][2], FBAIdbbox[3]+85]
                            else:
                                SKUbbox = [FBAIdbbox[0]-50, SingleSKUspan['bbox'][1], FBAIdbbox[2]+50, SingleSKUspan['bbox'][3]]

                            page.insert_textbox(SKUbbox, SKU, fontname='Helvetica-Bold', fontsize=SingleSKUspan['size'], rotate=rotation)
                            break
                if FBAId:
                    del FBAId
                    break

            # An error occured: line not found in block
            except:
                pass

    print(f"\tSaving to {save_location}")
    doc.save(save_location)
    return


def scrape_packlist(doc):
    """
    Scrapes relevant information of Amazon's shipment & box list for ST and SoStocked uploading
    Input: fitz doc
    Return: pandas df
    Note: Amazon started sending multiple shipments in one packlist
    """
    print("Scraping packing list")
    data = pd.DataFrame(columns=['Product Description', 'SKU', 'Quantity', 'PCS/Box', 'Boxes', 'Box Label #'])

    for page in doc:
        rows = page.get_text().split('\n')
        print(rows)

        # scraping though regrex & index positions
        for index, row in enumerate(rows):
            # Box label and weight (Box 4 of 4 - 46.30lb)
            if 'box' and 'lb' in row.lower():
                boxLabel  = re.findall(r'Box (\d+)', row, re.I)[0]
                boxWeight = re.findall(r'(\d+)\s?lb', row, re.I)[0]
            # Ship to FC location (SMF3)
            elif 'ship to:' in row.lower():
                FC_location = rows[index + 2]
                shipAddress = f"{rows[index+3]} {rows[index+4]}, {rows[index+5]}"
            # Shipment Name & Shipment Number 
            elif 'created:' in row.lower():
                dateCreated  = re.findall(r'Created: (.+) ', row, re.I)[0]
                shipmentName =  re.sub(r"[/:]", "-", rows[index - 1])
                print(f"Shipment Name IS: {shipmentName}")
                if 'shipment' in shipmentName: # multiple shipments (ST to AMZ Oct 2022 Shipment 2)
                    shipmentNumber  = re.findall(r'.+ (Shipment \d+)', span['text'])[0]
                else: # only 1 shipment (ST to AMZ Oct 2022)
                    shipmentNumber = 'Shipment 1'
            # FBA Box ID (FBA16XNXHNS9U000004)
            elif re.match(r'^FBA\w', row, re.IGNORECASE):
                fbaBoxIdNumber = row
                shipmentId     = fbaBoxIdNumber.split('U000')[0]
            # SKU, Quantity & other product details
            elif 'qty' in row.lower():
                quantity = int(re.findall(r'(\d+)', row)[0])
                sku = rows[index-1]
                # fixes sku
                if '...' in sku:
                    splittedSKU = sku.split('...')[0]
                    sku = masterData[masterData['SKU'].str.startswith(splittedSKU)]['SKU'].values[0]
                # other data from the master file
                rowData = masterData[masterData.SKU == sku]
                productDescription = rowData['Product Description'].values[0]
                unitsPerBox = int(rowData['Units per box'].values[0])

        # merging data
        if not any(data['SKU'].str.contains(sku)):
            data.loc[len(data)] = [productDescription, sku, quantity, unitsPerBox, 1, boxLabel]
        else:
            data.loc[data.SKU==sku, 'Quantity'] += quantity
            data.loc[data.SKU==sku, 'Boxes'] += 1
            data.loc[data.SKU==sku, 'Box Label #'] += f' - {boxLabel}'

    data[['Shipment Name', 'Shipment Number', 'Fulfillment Center', 
                                'Shipment ID', 'Shipping Address']] = shipmentName, shipmentNumber, FC_location, shipmentId, shipAddress
    return data


def summarize_packlists(directory):
    """Scrapes all shipping & box labels in a directory & summarizes it in a excel
    Input: full directory path
    Output: Excel shipping summary"""
    # scrapes data to pandas
    summary = pd.DataFrame()

    for filename in os.listdir(directory):
        filepath = os.path.join(directory, filename)
        # checks if shipping box label
        if filename.endswith('Box Labels.pdf'):
            with fitz.open(filepath) as doc:
                data = scrape_packlist(doc)
                summary = pd.concat([summary, data], axis=0, ignore_index=True)
    # writing & formatting to excel
    save_location = os.path.join(directory, 'Shipping Plan Summary.xlsx')
    # writer = pd.ExcelWriter(save_location)
    # summary.to_excel(writer, index=False, sheet_name='Summary')
    # workbook = writer.book
    # worksheet = writer.sheets['Summary']
    # worksheet.set_zoom(90)
    # # cell formats
    # header_format = workbook.add_format({'valign': 'vcenter',
    #                                         'align': 'center',
    #                                         'bold': True})
    # subheader_format = workbook.add_format({'bold': True,
    #                                         'bg_color': '#951f06',
    #                                         'font_color': '#FFFFFF'})
    # format = workbook.add_format({'align': 'center'})
    # columns = ['Product Description', 'SKU', 'Quantity', 'PCS/Box', 'Boxes', 'Box Label #']
    # data = summary[columns]
    # # writing column headers
    # for col_num, value in enumerate(summary.columns.values):
    #     worksheet.write(2, col_num, value, header_format)


    summary.to_excel(save_location, index=False)
    return


def shippingtree_orderimport(directory):
    """Creates Shipping Tree order template csv from the 'Shipping Plan Summary.xlsx'"""
    filepath = os.path.join(directory, 'Shipping Plan Summary.xlsx')
    summary = pd.read_excel(filepath)

    order_import = pd.DataFrame(columns=['sku', 'quantity'])
    for sku in summary['SKU'].unique():
        total_quantity = summary.loc[summary.SKU == sku, 'Quantity'].sum()
        order_import.loc[len(order_import)] = [sku, total_quantity]

    save_location = os.path.join(directory, 'order-import-template.csv')
    order_import.to_csv(save_location, index=False)
    return


def create_shippinguploads(file):
    """
    Scrape and add Amazon shipment & box labels
    Input: file name of the pdf file
    Return: shipping tree upload directory
    """
    with fitz.open(file) as doc:
        doc = fitz.open(file)
    # Scrapes relevant information
    data = scrape_packlist(doc)
    # Creates shipment folder
    shipmentDirectory = os.path.join(amazonShipmentsDirectory, data['Shipment Name'].values[0])
    print(f"Creating folder {shipmentDirectory}")
    os.makedirs(shipmentDirectory, exist_ok=True)
    # Fixes broken SKUs & save it to the shipment folder
    file_name = f"{data['Shipment Name'].values[0]} - {data['Shipment Number'].values[0]} Box Labels.pdf"
    save_location = os.path.join(shipmentDirectory, file_name)
    add_amazon_sku(doc, save_location)
    # Scrapes all box labels & creates shipping plan summary
    summarize_packlists(shipmentDirectory)
    # ST order-import-template
    shippingtree_orderimport(shipmentDirectory)
    # SoStocked template shipment uploads
    sostockedImportShipmentDirectory = sostocked_shipment(data, shipmentDirectory)
    return sostockedImportShipmentDirectory



if __name__ == '__main__':
    file = os.path.join(downloads_folder, 'package-FBA170YJ3X3G.pdf')
    create_shippinguploads(file)
    # with fitz.open(file) as doc:
    #     scrape_packlist(doc)
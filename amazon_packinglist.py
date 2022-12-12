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
shippingtreeDirectory = os.path.join(currentDirectory, 'Shipping Tree Uploads')

# Check for dump folders
os.makedirs(shippingtreeDirectory, exist_ok=True)

# Master Data File
masterData = pd.read_excel('Master Data File.xlsx')

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
    """
    print("Scraping packing list")
    data = pd.DataFrame(columns=['Product Description', 'SKU', 'Quantity', 'PCS/Box', 'Boxes', 'Box Label #'])

    for page in doc:
        # Skips shipment label in 'Thermal' format
        if "PLEASE LEAVE THIS LABEL UNCOVERED" not in page.get_text('json'):
            print("NOT IN PAGE COVER")
            continue
        pageText = json.loads(page.get_text('json'))
        pageBlocks = pageText['blocks']

        # Scrapes page > blocks > lines > span > span['texts]
        for block in pageBlocks:    # iterates through blocks, lines & spans
            try:
                for line in block['lines']:
                    for span in line['spans']:
                        # Box label and weight (Box 4 of 4 - 46.30lb)
                        if 'box' and 'lb' in span['text'].lower():
                            boxLabel  = re.findall(r'\d+', span['text'])[0]
                            boxWeight = span['text'].split('-')[1].strip()
                        # Ship to FC location (SMF3)
                        if 'ship to:' in span['text'].lower():
                            shipToBlockNumber = block['number'] + 2     # bottom of SHIP TO:
                            shipTo = pageBlocks[shipToBlockNumber]['lines'][0]['spans'][0]['text']
                            shipAddress1 = pageBlocks[shipToBlockNumber + 1]['lines'][0]['spans'][0]['text']
                            shipAddress2 = pageBlocks[shipToBlockNumber + 2]['lines'][0]['spans'][0]['text']
                            shipAddress3 = pageBlocks[shipToBlockNumber + 3]['lines'][0]['spans'][0]['text']
                            shipAddress = shipAddress1 + shipAddress2 + ', ' + shipAddress3 + f' ({shipTo})'
                        # Shipment Name & Shipment Number (ST to AMZ Oct 2022 Shipment 2)
                        if 'shipment' in span['text'].lower():
                            shipmentName = re.findall(r'(.+) Shipment', span['text'])[0]
                            shipmentNumber  = re.findall(r'.+ (Shipment \d+)', span['text'])[0]
                        # Date Created (Created: 2022/10/14 09:05 PDT (-07))
                        if 'created:' in span['text'].lower():
                            dateCreated = re.findall(r'Created: (.+) ', span['text'])[0]
                        # FBA Box ID  (FBA16XNXHNS9U000004), SKU & Quantity
                        if re.match(r'^FBA\w', span['text']):
                            fbaBoxIdNumber = span['text'] 
                            shipmentId     = fbaBoxIdNumber.split('U000')[0]
                            fbaBoxIdBlockNumber = block['number']
                            skuBlockNumber = fbaBoxIdBlockNumber + 2    # bottom of FBA box id
                            sku = pageBlocks[skuBlockNumber]['lines'][0]['spans'][0]['text']
                            # fixes sku
                            if '...' in sku:
                                splittedSKU = sku.split('...')[0]
                                sku = masterData[masterData['SKU'].str.startswith(splittedSKU)]['SKU'].values[0]
                            quantityBlockNumber = skuBlockNumber + 1
                            quantity = int(pageBlocks[quantityBlockNumber]['lines'][0]['spans'][0]['text'].split()[-1])
                            # other data from the master file
                            rowData = masterData[masterData.SKU == sku]
                            productDescription = rowData['Product Description'].values[0]
                            unitsPerBox = int(rowData['Units per box'].values[0])

            # An error occured: line not found in block
            except Exception as e:
                print(f"#ERROR! {e}")
                pass

        # merging data
        if not any(data['SKU'].str.contains(sku)):
            data.loc[len(data)] = [productDescription, sku, quantity, unitsPerBox, 1, boxLabel]
        else:
            data.loc[data.SKU==sku, 'Quantity'] += quantity
            data.loc[data.SKU==sku, 'Boxes'] += 1
            data.loc[data.SKU==sku, 'Box Label #'] += f' - {boxLabel}'

    metadata = pd.DataFrame({'Shipment Name': [shipmentName], 'Shipment Number': [shipmentNumber], 
                'Shipment ID': [shipmentId], 'Shipping Address': [shipAddress]})
    print(metadata)
    return data, metadata


def summarize_packlists(directory):
    """Scrapes all shipping & box labels in a directory & summarizes it in a excel
    Input: full directory path
    Output: Excel shipping summary"""
    summary = pd.DataFrame()

    for filename in os.listdir(directory):
        filepath = os.path.join(directory, filename)
        # checks if shipping box label
        if filename.endswith('Box Labels.pdf'):
            with fitz.open(filepath) as doc:
                data, metadata = scrape_packlist(doc)
                metadata = pd.concat([metadata] * len(data), ignore_index=True)
                metadata['SKU'] = data['SKU']
                mergedData = pd.merge(data, metadata, how='inner', on='SKU')
                summary = pd.concat([summary, mergedData], axis=0, ignore_index=True)

    save_location = os.path.join(directory, 'Shipping Plan Summary.xlsx')
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


def create_shipingplan(data, metadata):     # depracated
    title = f"{metadata['Shipment Name'].values[0]} - {metadata['Shipment Number'].values[0]}"
    fileName = f"{title} shipping plan.pdf"
    with PdfPages(fileName) as pdf:
        print(title)
        fig, axs = plt.subplots(figsize=(8.27, 11.69))
        fig.suptitle(title + '\n' + metadata['Shipment ID'].values[0] + '\n' + metadata['Shipping Address'].values[0])
        axs.axis('tight')
        axs.axis('off')
        table = axs.table(cellText=data.values, colLabels=data.columns, loc='upper center', fontsize=1)
        plt.show()
        pdf.savefig(fig, bbox_inches='tight')


def scrape_add_packinglist(file):
    """
    Scrape and add Amazon shipment & box labels
    Input: file name of the pdf file
    Return: shipping tree upload directory
    """
    with fitz.open(file) as doc:
        doc = fitz.open(file)
    # Scrapes relevant information
    data, metadata = scrape_packlist(doc)
    # Creates shipment folder
    shipmentDirectory = os.path.join(shippingtreeDirectory, metadata['Shipment Name'].values[0])
    print(f"Creating folder {shipmentDirectory}")
    os.makedirs(shipmentDirectory, exist_ok=True)
    # Fixes broken SKUs & save it to the shipment folder
    file_name = f"{metadata['Shipment Name'].values[0]} - {metadata['Shipment Number'].values[0]} Box Labels.pdf"
    save_location = os.path.join(shipmentDirectory, file_name)
    add_amazon_sku(doc, save_location)
    # Scrapes all box labels & creates shipping plan summary
    summarize_packlists(shipmentDirectory)
    # ST order-import-template
    shippingtree_orderimport(shipmentDirectory)
    return shipmentDirectory



if __name__ == '__main__':
    file = os.path.join(downloads_folder, 'package-FBA16VFPZNRP.pdf')
    scrape_add_packinglist(file)

    # with fitz.open(file) as doc:
    #     scrape_packlist(doc)1
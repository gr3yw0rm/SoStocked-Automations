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
import win32com.client as win32

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
    Return: 2 pandas df
    Note: Amazon started sending multiple shipments in one packlist
    """
    print("Scraping packing list")
    detailed_data = pd.DataFrame(columns=['Product Description', 'SKU', 'Qty', 'PCS/Box', 'Boxes', 'Box Label #', 
                                          'Shipment Name', 'Shipment Number', 'Fulfillment Center', 
                                          'Shipment ID', 'Shipping Address'])
    pdf_data = pd.DataFrame(columns=['Product Description', 'SKU', 'Qty', 'PCS/Box', 'Boxes', 'Box Label #', 'Shipment Name', 'Shipment Number'])

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
                if re.search(r'Shipment \d+', shipmentName, re.I): # multiple shipments (ST to AMZ Oct 2022 Shipment 2)
                    shipmentNumber  = re.findall(r'(Shipment \d+)', shipmentName, re.I)[0]
                    shipmentName = re.sub(r'Shipment \d+$', '', shipmentName).strip()   # removes shipment number at the end
                else: # only 1 shipment (ST to AMZ Oct 2022)
                    shipmentNumber = 'Shipment 1'
                print(f"Shipment Name IS: {shipmentName} with {shipmentNumber}")
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
    
        # adding row to detailed summary data
        page_data = pd.DataFrame([{'Product Description': productDescription, 'SKU': sku, 'Qty': quantity, 'PCS/Box': unitsPerBox, 'Boxes': 1, 
                     'Box Label #': boxLabel, 'Shipment Name': shipmentName, 'Shipment Number': shipmentNumber, 
                     'Fulfillment Center': FC_location, 'Shipment ID': shipmentId, 'Shipping Address': shipAddress}])
        detailed_data = pd.concat([detailed_data, page_data], ignore_index=True)
                
        # aggregating data for the PDF summary data
        if not any(pdf_data['SKU'].str.contains(sku)): # if exists
            pdf_data.loc[len(pdf_data)] = [productDescription, sku, quantity, unitsPerBox, 1, boxLabel, shipmentName, shipmentNumber]
        else:
            pdf_data.loc[pdf_data.SKU==sku, 'Qty'] += quantity
            pdf_data.loc[pdf_data.SKU==sku, 'Boxes'] += 1
            pdf_data.loc[pdf_data.SKU==sku, 'Box Label #'] += f' - {boxLabel}'
    return pdf_data, detailed_data


def summarize_packlists(directory):
    """Scrapes all shipping & box labels in a directory & summarizes it in a excel
    Input: full directory path
    Output: Excel & PDF shipping summary"""
    shipment_summary = pd.DataFrame()
    detailed_summary = pd.DataFrame()
    for filename in os.listdir(directory):
        filepath = os.path.join(directory, filename)
        # checks if shipping box label
        if filename.endswith('Box Labels.pdf'):
            print(filename)
            # appends to main df
            with fitz.open(filepath) as doc:
                shipment_summary_temp, detailed_summary_temp = scrape_packlist(doc)
                shipment_summary = pd.concat([shipment_summary, shipment_summary_temp], ignore_index=True)
                detailed_summary = pd.concat([detailed_summary, detailed_summary_temp], ignore_index=True)

    # writing & formatting to excel
    save_location = os.path.join(directory, 'Shipping Plan Summary.xlsx')
    with pd.ExcelWriter(save_location, engine='xlsxwriter') as writer:
        pd.DataFrame().to_excel(writer, index=False, sheet_name='Summary') # blank sheet
        workbook = writer.book
        worksheet = writer.sheets['Summary']
        worksheet.set_zoom(100)
        worksheet.set_margins(0.15, 0.15, 0.75, 0.75)
        left_format   = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True})
        center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        worksheet.set_column('A:A', 33, left_format)
        worksheet.set_column('B:B', 33, left_format)
        worksheet.set_column('C:C', 4, center_format)
        worksheet.set_column('D:D', 7, center_format)
        worksheet.set_column('E:E', 5, center_format)
        worksheet.set_column('F:F', 17, left_format)
        # writing title
        merged_title_format = workbook.add_format({'font_size': 20, 'bold': 1, 'align': 'center', 'valign': 'vcenter'})
        worksheet.merge_range('A1:F1', "Shipping Plan Summary", merged_title_format)
        # writing summary per shipment number
        start_row = 2
        for shipment_no in shipment_summary['Shipment Number'].unique():
            shipmentDataDetailed = shipment_summary[shipment_summary['Shipment Number'] == shipment_no]
            # writing shipment name above table
            shipmentName        = shipmentDataDetailed['Shipment Name'].values[0]
            shipmentNumber      = shipmentDataDetailed['Shipment Number'].values[0]
            shipmentTitle       = shipmentName + " " + shipmentNumber
            shipmentName_cell   = xl_rowcol_to_cell(start_row, 0)
            shipmentName_format = workbook.add_format({'font_size': 14, 'valign': 'vcenter', 'align': 'left', 'bold': True})
            worksheet.write(shipmentName_cell, shipmentTitle, shipmentName_format)
            start_row += 1
            # writing column headers
            pdf_cols = ['Product Description', 'SKU', 'Qty', 'PCS/Box', 'Boxes', 'Box Label #']
            shipment_data = shipmentDataDetailed[pdf_cols]
            header_format = workbook.add_format({'bold': True, 'bg_color': '#951f06', 'font_color': '#FFFFFF'})
            for col_num, value in enumerate(shipment_data.columns.values):
                cell = xl_rowcol_to_cell(start_row, col_num)
                worksheet.write(cell, value, header_format)
            start_row += 1
            # writing shipment data
            shipment_data.to_excel(writer, startrow=start_row, index=False, header=False, sheet_name='Summary')
            start_row += shipment_data.shape[0] + 1
        # detailed summary on sheet2
        detailed_summary.to_excel(writer, index=False, sheet_name='Detailed Summary')
        worksheet2 = writer.sheets['Detailed Summary']
        worksheet2.set_zoom(140)
        worksheet2.freeze_panes(1, 0)
        worksheet2.autofit()

    # converting shipment summary to pdf
    pdf_save_location = os.path.join(directory, 'Shipping Plan Summary.pdf')
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(save_location)
    wb.ActiveSheet.ExportAsFixedFormat(0, pdf_save_location)
    wb.Close()
    excel.Quit()
    return


def shippingtree_orderimport(directory):
    """Creates Shipping Tree order template csv from the 'Shipping Plan Summary.xlsx'"""
    filepath = os.path.join(directory, 'Shipping Plan Summary.xlsx')
    summary = pd.read_excel(filepath, sheet_name='Detailed Summary')

    order_import = pd.DataFrame(columns=['sku', 'quantity'])
    for sku in summary['SKU'].unique():
        total_quantity = summary.loc[summary.SKU == sku, 'Qty'].sum()
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
    pdf_data, detailed_data = scrape_packlist(doc)
    # Creates shipment folder
    shipmentDirectory = os.path.join(amazonShipmentsDirectory, detailed_data['Shipment Name'].values[0])
    print(f"Creating folder {shipmentDirectory}")
    os.makedirs(shipmentDirectory, exist_ok=True)
    # Fixes broken SKUs & save it to the shipment folder
    file_name = f"{detailed_data['Shipment Name'].values[0]} - {detailed_data['Shipment Number'].values[0]} Box Labels.pdf"
    save_location = os.path.join(shipmentDirectory, file_name)
    add_amazon_sku(doc, save_location)
    # Scrapes all box labels & creates shipping plan summary
    print(shipmentDirectory)
    summarize_packlists(shipmentDirectory)
    # ST order-import-template
    shippingtree_orderimport(shipmentDirectory)
    # SoStocked template shipment uploads
    sostockedImportShipmentDirectory = sostocked_shipment(shipmentDirectory)
    return sostockedImportShipmentDirectory


if __name__ == '__main__':
    file = os.path.join(downloads_folder, 'package-FBA16XNY1B51 (1).pdf')
    create_shippinguploads(file)
    # with fitz.open(file) as doc:
    #     doc = fitz.open(file)
    # # Scrapes relevant information
    # data = scrape_packlist(doc)
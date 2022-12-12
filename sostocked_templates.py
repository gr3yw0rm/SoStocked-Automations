import os
import sys
import glob
import pathlib
import shutil
import pandas as pd
import numpy as np
import datetime as dt


# determine if application is a script file or frozen exe
if getattr(sys, 'frozen', False):
    currentDirectory = os.path.dirname(os.path.realpath(sys.executable))
    print(currentDirectory)
elif __file__:
    currentDirectory = os.path.dirname(__file__)

# Folder locations
os.chdir(currentDirectory)
downloadsDirectory = os.path.join(os.getenv('HOMEPATH'), 'Downloads')
sostockedDirectory = os.path.join(currentDirectory, 'SoStocked Import Warehouse Inventories')
sostockedShipmentsDirectory = os.path.join(currentDirectory, 'SoStocked Bulk Import Shipments')
amazonManifestsDirectory = os.path.join(currentDirectory, 'Amazon Manifest Workflows')

# Checks for dump folders
os.makedirs(sostockedDirectory, exist_ok=True)
os.makedirs(sostockedShipmentsDirectory, exist_ok=True)
os.makedirs(amazonManifestsDirectory, exist_ok=True)


# Master data file
masterDataFile = os.path.join(currentDirectory, 'Master Data File.xlsx')
activeProducts = pd.read_excel(masterDataFile, sheet_name='All Products')
activeProducts = activeProducts.loc[activeProducts['Status'] == 'Active']

def update_inventory(file='latest'):
    if file == 'latest':
        # Finds latest Shipping Tree's warehouse inventory report (Shopify > Analytics > Reports > Units Sold)
        inventorySales_files = glob.glob(os.path.join(downloadsDirectory, 'inventory_sales*.csv'))
        if not inventorySales_files:
            print("Shopify Inventory Sales not found in your Downloads folder.")

        file = max(inventorySales_files, key=os.path.getctime)
        print(f"Found {file}")

    # Read files to dataframe
    inventorySales = pd.read_csv(file)
    inventorySales = inventorySales[inventorySales['ending_quantity'] >= 0]     # dropping negative quantity i.e. charities
    template_location = os.path.join(currentDirectory, 'Templates', 'SoStocked-WH-Inventory-Import-Template.xlsx')
    sostockedTemplate = pd.read_excel(template_location)
    vendorData        = pd.read_excel('Master Data File.xlsx', sheet_name='Vendors')

    # Combines SoStocked's active products and ST/Shopify's inventory levels and vendor data
    stInventory = inventorySales.merge(activeProducts, 'left', left_on='product_variant_sku', right_on='SKU')
    stInventory['Vendor Name'] = 'Shipping Tree'
    stInventory['Vendor ID']   = vendorData.loc[vendorData['Vendor Name***'] == 'Shipping Tree', 'Vendor ID - SoStocked'].values[0]

    # Inputs to SoStocked's template
    template_cols    = ['Vendor ID - SoStocked', 'Vendor Name (aka warehouse name)***', 
                        'Quantity*** (in units)', 'Product Name', 'ASIN', 'SKU', 'Product ID - SoStocked']
    stInventory_cols = ['Vendor ID', 'Vendor Name', 'ending_quantity', 'Product Description', 
                        'ASIN', 'product_variant_sku', 'Product ID - SoStocked']
    sostockedTemplate[template_cols] = stInventory[stInventory_cols]

    # Creates new & writes to template
    datetime = dt.datetime.today().strftime('%b_%d_%Y-%I_%M%p')
    uploadTemplate = f'SoStocked-WH-Inventory-Import-{datetime}.xlsx'
    uploadTemplate_loc = os.path.join(sostockedDirectory, uploadTemplate)
    shutil.copy(template_location, uploadTemplate_loc)

    with pd.ExcelWriter(uploadTemplate_loc, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        sostockedTemplate.to_excel(writer,sheet_name='Warehouse Inventory levels', index=False)
        writer.save()

    # Moves Shopify Inventory Report to dump folder
    # shutil.move(file, shopifyDirectory)
    print(str(uploadTemplate_loc))
    return str(uploadTemplate_loc)


def send_to_amazon(file=None):
    # if file is None:
    #     file = os.path.join(downloadsDirectory, 'Nora-s-Nursery-Inc--Product-Calculations-Download-  -4436.xlsx')

    forecast = pd.read_excel(file)[['SKU', 'TRANSFER', 'Units per Carton (Case)', 'Transfer Case Qty']]

    workflowCols = ['SKU', 'Box length (in)', 'Box width (in)', 'Box height (in)', 'Box weight (lb)']
    workflowTransfers = forecast.merge(activeProducts[workflowCols], how='left', on='SKU')

    # Adding empty cols to match template
    workflowTransfers.insert(2, 'Prep Owner', np.nan)
    workflowTransfers.insert(3, 'Labeling owner', np.nan)
    workflowTransfers.insert(4, 'Expiration Date (MM/DD/YYYY)', np.nan)

    # Creates new & writes to template
    datetime = dt.datetime.today().strftime('%b_%d_%Y-%I_%M%p')
    workflowTemplate = f'Amazon Manifest Workflows/Manifest Workflow Template_{datetime}.xlsx'
    template_location = os.path.join(currentDirectory, 'Templates', 'Manifest Workflow Template.xlsx')
    shutil.copy(template_location, workflowTemplate)

    with pd.ExcelWriter(workflowTemplate, mode='a', if_sheet_exists='overlay') as writer:
        workflowTransfers.to_excel(writer, startrow=6, header=False, index=False, sheet_name='Create workflow – template')

    # # SoStocked Import Shipment Template
    # forecast = forecast.iloc[:, :-2]
    # shipmentCols = ['ASIN', 'SKU']
    # shipmentTransfers = forecast.merge(activeProducts[shipmentCols], how='left', on='SKU')
    # shipmentTransfers[['ASIN Marketplace', 'SKU Marketplace', 'FN SKU Marketplace']] = 'US'
    # shipmentTransfers[['Units Arrived', 'Cost Per Unit']] = 0
    # shipmentTransfers['FN SKU'] = np.nan
    # shipmentTransfers = shipmentTransfers[['ASIN Marketplace', 'ASIN', 'SKU Marketplace', 'SKU', 
    #                                         'FN SKU Marketplace', 'FN SKU', 'TRANSFER', 'Units Arrived', 'Cost Per Unit']]

    # shipmentTemplate = f'{sostockedShipmentsDirectory}/SoStocked Bulk Import Shipment Template_{datetime}.xlsx'
    # shutil.copy('SoStocked-Bulk-Import-Shipment-Template.xlsx', shipmentTemplate)

    # with pd.ExcelWriter(shipmentTemplate, mode='a', if_sheet_exists='overlay') as writer:
    #     shipmentTransfers.to_excel(writer, startrow=1, header=False, index=False, sheet_name='Edit Shipment Import Export')

    return workflowTemplate
        

def split_shipment(file):
    if not file:
        file = os.path.join(downloadsDirectory, 'FBA16RTT5ZY7.csv')

    # Getting number of rows of shitment details
    with open(file) as f:
        readFile = f.read()
        nRows_details = readFile.split('\n\n')[0].count('\n')
        nRows_sku     = readFile.split('\n"SKU"')[0].count('\n') + 1

    shipmentDetails = pd.read_csv(file, delimiter=',', nrows=nRows_details, index_col=0).T
    shipmentDetails['Shipment name'] = shipmentDetails['Shipment name'].str.replace(r'\W', ' ')
    packList = pd.read_csv(file, delimiter=',', skiprows=nRows_sku)

    # Creates new & writes to template
    sostockedImportShipment = os.path.join(sostockedShipmentsDirectory, f"{shipmentDetails['Shipment name'][0]} - {shipmentDetails['Shipment ID'][0]}.xlsx")
    template_location = os.path.join(currentDirectory, 'Templates', 'SoStocked-Bulk-Import-Shipment-Template.xlsx')
    shutil.copy(template_location, sostockedImportShipment)
    
    packList[['ASIN Marketplace', 'SKU Marketplace', 'FN SKU Marketplace']] = 'US'
    packList[['Units Arrived', 'Cost Per Unit']] = 0
    packList = packList[['ASIN Marketplace', 'ASIN', 'SKU Marketplace', 'SKU', 'FN SKU Marketplace', 'FNSKU', 'Total units', 'Units Arrived', 'Cost Per Unit']]

    with pd.ExcelWriter(sostockedImportShipment, mode='a', if_sheet_exists='overlay') as writer:
        packList.to_excel(writer, startrow=1, header=False, index=False, sheet_name='Edit Shipment Import Export')
    
    return sostockedImportShipment

if __name__ == '__main__':
    # print(os.getcwd())
    pass
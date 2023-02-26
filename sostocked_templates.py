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
amazonManifestsDirectory = os.path.join(currentDirectory, 'Amazon Manifest Workflows')

# Checks for dump folders
os.makedirs(sostockedDirectory, exist_ok=True)
os.makedirs(amazonManifestsDirectory, exist_ok=True)

# Master data file
masterDataFile = os.path.join(currentDirectory, 'Master Data File.xlsx')
activeProducts = pd.read_excel(masterDataFile, sheet_name='All Products').dropna(how='all').fillna('')
activeProducts = activeProducts.loc[activeProducts['Status'] == 'Active']

# Template locations
shipmentTemplate_location = os.path.join(currentDirectory, 'Templates', 'SoStocked-Bulk-Import-Shipment-Template.xlsx')


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
    formatted_uploadTemplate_loc = os.sep.join(os.path.normpath(uploadTemplate_loc).split(os.sep)[-2:]).replace(os.sep, '>>')
    return str(formatted_uploadTemplate_loc)


def send_to_amazon(file=None):
    forecast = pd.read_excel(file)
    # Dropping Shopfy & other cols
    columns = ['SKU', 'TRANSFER', 'Units per Carton (Case)', 'Transfer Case Qty']
    forecast = forecast.loc[(forecast.Marketplace != 'Shopfy') & (forecast['TRANSFER'] > 0), columns]

    workflowCols = ['SKU', 'Box length (in)', 'Box width (in)', 'Box height (in)', 'Box weight (lb)']
    workflowTransfers = forecast.merge(activeProducts[workflowCols], how='left', on='SKU')

    # Adding empty cols to match template
    workflowTransfers.insert(2, 'Prep Owner', np.nan)
    workflowTransfers.insert(3, 'Labeling owner', np.nan)
    workflowTransfers.insert(4, 'Expiration Date (MM/DD/YYYY)', np.nan)

    # Creates new & writes to template
    datetime = dt.datetime.today().strftime('%b_%d_%Y-%I_%M%p')
    workflowTemplate = os.path.join('Amazon Manifest Workflows', f'Manifest Workflow Template_{datetime}.xlsx')
    template_location = os.path.join(currentDirectory, 'Templates', 'Manifest Workflow Template.xlsx')
    shutil.copy(template_location, workflowTemplate)

    with pd.ExcelWriter(workflowTemplate, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
        workflowTransfers.to_excel(writer, startrow=6, header=False, index=False, sheet_name='Create workflow – template')

    formatted_workflowTemplate = os.sep.join(os.path.normpath(workflowTemplate).split(os.sep)[-2:]).replace(os.sep, '>>')
    return formatted_workflowTemplate


def sostocked_shipment(directory):
    """Converts scraped data from Amazon box labels into SoStocked Shipment template.
    Then saves it into a assigned shipping directory.
    Input: Amazon shipment directory
    Output: SoStocked upload shipment template"""
    shippingPlanSummaryPath = os.path.join(directory, 'Shipping Plan Summary.xlsx')
    detailed_data = pd.read_excel(shippingPlanSummaryPath, sheet_name='Detailed Summary')
    # groups data by SKU & FC location
    grouped_data = detailed_data.groupby(['SKU', 'Fulfillment Center']).sum().reset_index()
       
    # Creates new & writes to template for each FC
    for fc in grouped_data['Fulfillment Center'].unique():
        cleanData = pd.DataFrame(columns=['ASIN Marketplace', 'ASIN', 'SKU Marketplace', 'SKU', 'FN SKU Marketplace', 
                                                'FNSKU', 'Quantity', 'Units Arrived', 'Cost Per Unit'])
        cleanData[['SKU', 'Quantity']] = grouped_data.loc[grouped_data['Fulfillment Center'] == fc, ['SKU', 'Qty']]
        # extracting & combining shipment numbers 1,2,3
        shipment_numbers = detailed_data.loc[detailed_data['Fulfillment Center'].str.contains(fc), 'Shipment Number']
        shipment_numbers = ",".join(set(shipment_numbers.str.extract(r'(\d+)').astype(str)[0].tolist()))
        file_name = f"SoStocked Import Shipment {shipment_numbers} - {fc}.xlsx"
        # moves sostocked blank shipment template to amazon shipment directory
        sostockedImportShipmentDirectory = os.path.join(directory, file_name)
        shutil.copy(shipmentTemplate_location, sostockedImportShipmentDirectory)
        print(cleanData)
        with pd.ExcelWriter(sostockedImportShipmentDirectory, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
            cleanData.to_excel(writer, startrow=1, header=False, index=False, sheet_name='Edit Shipment Import Export')

    formatted_sostockedImportShipmentDirectory = os.sep.join(os.path.normpath(sostockedImportShipmentDirectory).split(os.sep)[-3:-1]).replace(os.sep, '>>')
    return formatted_sostockedImportShipmentDirectory


if __name__ == '__main__':
    file = os.path.join(downloadsDirectory, 'Nora-s-Nursery-Inc--Product-Calculations-Download-20230119082749-6300 (2).xlsx')
    update_inventory(file)
    # directory = os.path.join(os.getcwd(), 'Amazon Shipments', 'ST to AMZ Oct 2022')
    # sostocked_shipment(directory)
    pass
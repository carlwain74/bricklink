import json
import argparse
import sys
import logging
from os import stat
from os.path import exists
from bricklink_api.auth import oauth
from bricklink_api.catalog_item import get_price_guide, get_item, get_item_image, Type, NewOrUsed
from bricklink_api.category import get_category
import html
from html.parser import HTMLParser
import xlsxwriter

logging.basicConfig(
    format='%(asctime)s %(levelname)-8s %(message)s',
    level=logging.INFO,
    datefmt='%Y-%m-%d %H:%M:%S')

# Get info and return key elements
def getDetails(set_number):
    logging.debug("Getting details for " + str(set_number))
    h = html.parser

    if set_number == "40158":
        itemType = Type.GEAR
    else:
        itemType = Type.SET
    
    json_obj = get_price_guide(itemType, set_number, new_or_used=NewOrUsed.NEW, country_code="US", region="north_america", auth=auth)

    logging.debug(json.dumps(json_obj, indent=4, sort_keys=True))
    meta = json_obj['meta']

    if meta['code'] == 200:
        data = json_obj['data']
        #print("Core Data")
        #print(data)
        #print("Meta Data")
        #print(meta)

        typeData = get_item(itemType, set_number, auth=auth)      
        logging.debug(json.dumps(typeData, indent=4, sort_keys=True))
        
        categoryObj = get_category(typeData['data']['category_id'], auth=auth)
        logging.debug(json.dumps(categoryObj, indent=4, sort_keys=True))
        
        elem_data = {}
        elem_data[set_number] = {}
        elem_data[set_number]['name'] = h.unescape(typeData['data']['name'])
        elem_data[set_number]['category'] = h.unescape(categoryObj['data']['category_name'])
        elem_data[set_number]['avg'] = round(int(float(data['avg_price'])))
        elem_data[set_number]['max'] = round(int(float(data['max_price'])))
        elem_data[set_number]['min'] = round(int(float(data['min_price'])))
        elem_data[set_number]['quantity'] = data['unit_quantity']
        elem_data[set_number]['currency'] = data['currency_code']

        return elem_data
    else:
        logging.warning("API Error!! " + str(meta['code']))
        logging.warning("API Message!! " + str(meta['message']))
        return 0

def printDetails(elementData, number):
        logging.info("Item: " + number)
        logging.info("  Name: " + elementData['name'])
        logging.info("  Category: " + elementData['category'])
        logging.info("  Avg Price: " + str(elementData['avg']) + " " + elementData['currency'])
        logging.info("  Max Price: " + str(elementData['max']) + " " + elementData['currency'])
        logging.info("  Min Price: " + str(elementData['min']) + " " + elementData['currency'])
        logging.info("  Quantity avail: " + str(elementData['quantity']))

# fill in with your data from https://www.bricklink.com/v2/api/register_consumer.page
consumer_key = ""
consumer_secret = ""
token_value = ""
token_secret = ""
auth = oauth(consumer_key, consumer_secret, token_value, token_secret)

parser = argparse.ArgumentParser()
parser.add_argument('-s', '--set', type=str)
parser.add_argument('-f', '--file', type=str)
parser.add_argument('-v', '--verbose', action="store_true")
args = parser.parse_args()

set_num = args.set
filename = args.file

workbook = xlsxwriter.Workbook('Expenses01.xlsx')
worksheet = workbook.add_worksheet('For Sale')

# Start from the first cell. Rows and columns are zero indexed.
row = 1
col = 1

worksheet.set_column('B:B', 20)
worksheet.set_column('C:C', 20)
worksheet.set_column('D:D', 20)
worksheet.set_column('E:E', 20)
worksheet.set_column('F:F', 20)
worksheet.set_column('G:G', 20)
worksheet.set_column('F:F', 20)

cell_format = workbook.add_format()
cell_format.set_align('center')
cell_format.set_align('vcenter')

header_format = workbook.add_format()
header_format.set_align('center')
header_format.set_align('vcenter')
header_format.set_bold()
header_format.set_bg_color('#C0C0C0')

xls_headers = ['Item', 'Name', 'Category', 'Avg Price', 'Min Price', 'Max Price', 'Quantity']

i=0
for x in xls_headers:
    worksheet.write(row, col+i, x, header_format)
    i += 1

if args.verbose:
    logging.getLogger().setLevel(logging.DEBUG)

if set_num:
    res = getDetails(set_num)
    logging.debug(json.dumps(res, indent=4, sort_keys=True))
    for key in res:
        printDetails(res[key],key)
elif filename:
    if exists(filename):
        logging.info("Processing sets in " + filename)

        if stat(filename).st_size == 0:
            logging.error("File is empty!!")
            sys.exit()
        else:
            FileHandler = open(filename, "r")
            total = 0
            while True:
                line = FileHandler.readline()
                if not line:
                    break
                #print(line.strip())
                number = line.strip()
                res = getDetails(number)
                for key in res:
                    printDetails(res[key],key)
                    logging.debug(json.dumps(res, indent=4, sort_keys=True))
                    total += res[key]['avg']

                    row += 1
                    worksheet.write(row,col, key, cell_format)
                    worksheet.write(row,col+1, res[key]['name'], cell_format)
                    worksheet.write(row,col+2, res[key]['category'], cell_format)
                    worksheet.write(row,col+3, res[key]['avg'], cell_format)
                    worksheet.write(row,col+4, res[key]['min'], cell_format)
                    worksheet.write(row,col+5, res[key]['max'], cell_format)
                    worksheet.write(row,col+6, res[key]['quantity'], cell_format)

            logging.info("Total: " + str(total) + "USD")

workbook.close()

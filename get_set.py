"""
This is neat little utility that gets pricing for Lego sets.
"""
import json
import argparse
import sys
import logging
from os import stat
from os.path import exists
#from bricklink_api.auth import oauth
#from bricklink_api.catalog_item import get_price_guide, get_item, get_item_image, Type, NewOrUsed
#from bricklink_api.category import get_category

from bricklink_py import Bricklink


import html
from html.parser import HTMLParser
import xlsxwriter
import configparser
from datetime import datetime

logging.basicConfig(
    format='%(asctime)s %(levelname)-8s %(message)s',
    level=logging.INFO,
    datefmt='%Y-%m-%d %H:%M:%S')

"""
This calls the API functions to get the data.
"""
def getDetails(session, set_number):
    logging.debug("Getting details for " + str(set_number))
    h_parse = html.parser

    if set_number == "40158":
        item_type = "GEAR"
    else:
        item_type = "SET"

    try:
        current_items = session.catalog_item.get_price_guide(item_type, set_number, new_or_used="N",
                                                             country_code="US", region="north_america")

        past_sales = session.catalog_item.get_price_guide(item_type, set_number, new_or_used="N", \
                                                          guide_type="sold", country_code="US", region="north_america")
    except Exception as e:
        logging.exception("Failed to get price guide for item" + str(e))
        sys.exit()

    logging.debug(json.dumps(current_items, indent=4, sort_keys=True))
    logging.debug(json.dumps(past_sales, indent=4, sort_keys=True))

    # 

    type_data = session.catalog_item.get_item(item_type, set_number)

    logging.debug(json.dumps(type_data, indent=4, sort_keys=True))

    category_data = session.category.get_category(type_data['category_id'])
    logging.debug(json.dumps(category_data, indent=4, sort_keys=True))

    elem_data = {}
    elem_data[set_number] = {}
    elem_data[set_number]['name'] = h_parse.unescape(type_data['name'])
    elem_data[set_number]['category'] = h_parse.unescape(category_data['category_name'])
    elem_data[set_number]['avg'] = round(int(float(current_items['avg_price'])))
    elem_data[set_number]['max'] = round(int(float(current_items['max_price'])))
    elem_data[set_number]['min'] = round(int(float(current_items['min_price'])))
    elem_data[set_number]['quantity'] = current_items['unit_quantity']
    elem_data[set_number]['currency'] = current_items['currency_code']
    elem_data[set_number]['year'] = type_data['year_released']

    return elem_data


"""
This prints stuff to the screen.
"""
def print_details(element_data, number):
    logging.info("Item: " + number)
    logging.info("  Name: " + element_data['name'])
    logging.info("  Category: " + element_data['category'])
    logging.info("  Avg Price: " + str(element_data['avg']) + " " + element_data['currency'])
    logging.info("  Max Price: " + str(element_data['max']) + " " + element_data['currency'])
    logging.info("  Min Price: " + str(element_data['min']) + " " + element_data['currency'])
    logging.info("  Quantity avail: " + str(element_data['quantity']))

"""
Setup XLS
"""
def setup_xls_writer():
    workbook = xlsxwriter.Workbook('Dec2023-combined.xlsx')

    now = datetime.now() # current date and time
    date_stamp = now.strftime("%m_%d_%Y")
    worksheet = workbook.add_worksheet('Items_'+date_stamp)

    # Start from the first cell. Rows and columns are zero indexed.
    row = 1
    col = 1

    worksheet.set_column('B:B', 20)
    worksheet.set_column('C:C', 30)
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

    xls_headers = ['Item', 'Name', 'Category', 'Avg Price', 'Min Price', 'Max Price', 'Quantity', 'Year']

    col_adjust = 0
    for headers in xls_headers:
        worksheet.write(row, col+col_adjust, headers, header_format)
        col_adjust += 1

    return workbook, worksheet, cell_format
"""
The main routine.
"""
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('-s', '--set', type=str)
    parser.add_argument('-f', '--file', type=str)
    parser.add_argument('-v', '--verbose', action="store_true")
    args = parser.parse_args()

    set_num = args.set
    filename = args.file

    config = configparser.ConfigParser()
    config.read('config.ini')

    # fill in with your data from https://www.bricklink.com/v2/api/register_consumer.page
    consumer_key = config['secrets']['consumer_key']
    consumer_secret = config['secrets']['consumer_secret']
    token_value = config['secrets']['token_value']
    token_secret = config['secrets']['token_secret']
    try:
        session = Bricklink(
            consumer_key=consumer_key,
            consumer_secret=consumer_secret,
            token=token_value,
            token_secret=token_secret
        )
    except Exception as e:
        logging.error('Could not get auth token')
        sys.exit(1)

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    if set_num:
        res = getDetails(session, set_num)
        if not res:
            sys.exit(1)
        logging.debug(json.dumps(res, indent=4, sort_keys=True))
        for key in res:
            print_details(res[key], key)
    elif filename:
        (workbook, worksheet, cell_format) = setup_xls_writer()

        if exists(filename):
            logging.info("Processing sets in " + filename)

            if stat(filename).st_size == 0:
                logging.error("File is empty!!")
                sys.exit()
            else:
                file_handler = open(filename, "r")
                total = 0
                row = 1
                col = 1
                while True:
                    line = file_handler.readline()
                    if not line:
                        break
                    #print(line.strip())
                    number = line.strip()
                    res = getDetails(session, number)
                    if not res:
                        sys.exit(1)
                    for key in res:
                        print_details(res[key], key)
                        logging.debug(json.dumps(res, indent=4, sort_keys=True))
                        total += res[key]['avg']

                        row += 1
                        worksheet.write(row, col, key, cell_format)
                        worksheet.write(row, col+1, res[key]['name'], cell_format)
                        worksheet.write(row, col+2, res[key]['category'], cell_format)
                        worksheet.write(row, col+3, res[key]['avg'], cell_format)
                        worksheet.write(row, col+4, res[key]['min'], cell_format)
                        worksheet.write(row, col+5, res[key]['max'], cell_format)
                        worksheet.write(row, col+6, res[key]['quantity'], cell_format)
                        worksheet.write(row, col+7, res[key]['year'], cell_format)


                logging.info("Total: " + str(total) + "USD")

        workbook.close()

if __name__ == '__main__':
    main()

"""
This is neat little utility that updates inventory on Bricklink.
"""
import json
import argparse
import sys
import logging
import os
from os import stat
from os.path import exists
from bricklink_api.auth import oauth
from bricklink_api.catalog_item import get_price_guide, get_item, get_item_image, Type, NewOrUsed
from bricklink_api.category import get_category
from bricklink_api.user_inventory import create_inventory, update_inventory
from bricklink_api.color import get_color_list
import html
from html.parser import HTMLParser
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment,Font,PatternFill
import configparser
from datetime import datetime


logging.basicConfig(
    format='%(asctime)s %(levelname)-8s %(message)s',
    level=logging.INFO,
    datefmt='%Y-%m-%d %H:%M:%S')

"""
Open inventory workbook
"""
def setup_xls_writer(xls_filename):
    try:
        if os.path.isfile(xls_filename) and os.access(xls_filename, os.R_OK):
            logging.info('Load excel file')
            workbook = load_workbook(filename=xls_filename)
    except Exception as exception:
        logging.error('Could not load excel file!' + str(exception))
        sys.exit(1)

    return workbook

def main():

    parser = argparse.ArgumentParser()
    parser.add_argument('-v', '--verbose', action="store_true")
    args = parser.parse_args()

    logging.info('Read configuration')
    config = configparser.ConfigParser()
    config.read('config.ini')

    # fill in with your data from https://www.bricklink.com/v2/api/register_consumer.page
    consumer_key = config['secrets']['consumer_key']
    consumer_secret = config['secrets']['consumer_secret']
    token_value = config['secrets']['token_value']
    token_secret = config['secrets']['token_secret']

    try:
        auth = oauth(consumer_key, consumer_secret, token_value, token_secret)
    except Exception as error:
        logging.error('Could not get auth token' + str(error))
        sys.exit(1)

    workbook = setup_xls_writer('LegoParts.xlsx')

    worksheet = workbook['Sheet1']
    index = 4
    while worksheet.cell(row=index, column=3).value is not None:
        # Check if there is an inventory id
        if worksheet.cell(row=index, column=2).value is not None:
            inventory_id = worksheet.cell(row=index, column=2).value
            logging.info("Inventory Id: " + str(inventory_id))
        else:
            inventory_id = 0
        item_type = worksheet.cell(row=index, column=3).value
        logging.info("Item Type: " + str(item_type))
        item_num = worksheet.cell(row=index, column=4).value
        logging.info("Item Num: " + str(item_num))
        color = worksheet.cell(row=index, column=6).value
        logging.info("Color: " + str(color))
        price = worksheet.cell(row=index, column=7).value
        logging.info("Price: " + str(price))
        quantity = worksheet.cell(row=index, column=8).value
        logging.info("Quantity: " + str(quantity))
        condition = worksheet.cell(row=index, column=9).value
        logging.info("Condition: " + str(condition))
        description = worksheet.cell(row=index, column=12).value
        logging.info("Description: " + str(description))
        stockroom = worksheet.cell(row=index, column=16).value
        logging.info("Stockroom: " + str(stockroom))
        stockroom_id = worksheet.cell(row=index, column=17).value
        logging.info("Stockroom Id: " + str(stockroom_id))
        retain = worksheet.cell(row=index, column=18).value
        logging.info("Stockroom Id: " + str(retain))

        # Call Bricklink API
        inventory_item = {}
        inventory_item['item'] = {}
        inventory_item['item']['type'] = 'PART'
        inventory_item['item']['no'] = item_num
        inventory_item['color_id'] = '1'
        inventory_item['unit_price'] = price
        inventory_item['quantity'] = quantity
        inventory_item['new_or_used'] = condition
        inventory_item['description'] = description
        inventory_item['is_stock_room'] = 'true'
        inventory_item['stock_room_id'] = stockroom_id
        inventory_item['is_retain'] = 'true'
        logging.debug(inventory_item)

        if inventory_id == 0:
            logging.info('Creating Inventory Item')
            response = create_inventory(inventory_item, auth=auth)
            inventory_id = response['data']['inventory_id']
        else:
            logging.info('Updating Inventory Item')
            response = update_inventory(inventory_id, inventory_item, auth=auth)

        logging.debug(response)
        worksheet.cell(row=index, column=2).value = inventory_id

        index += 1
    workbook.save(filename='LegoParts.xlsx')


if __name__ == '__main__':
    main()
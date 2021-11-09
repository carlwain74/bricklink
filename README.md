# bricklink
Utilities for collecting data from Bricklink using their API

## Requirements
- Python3

## Packages
- XlsWriter (Used by get_set.py)
- openpyxl (Used by get_set_openpyxl.py)
- bricklink-api (Common)
- virtualenv

## Setup Virtual Environment

### Creat virtual environment 
```
python3 -m venv env
source env/bin/activate
```

Setup virtual env based on which version of Xls package you wish to make use of.

### XlsWriter
```
python3 -m pip install -r requirements_xlswriter.txt
```
### OpenPyXl
```
python3 -m pip install -r requirements_openpyxl.txt
```

## Configuration File

The file `config.ini` needs to be populated.

Head over to https://www.bricklink.com/v2/api/register_consumer.page where you can create a bricklink account and setup your API access.

Populate the file as follows after the the Access Token has been created for the allowedIP.
```
[secrets]
consumer_key = <ConsumerKey>
consumer_secret = <ConsumerSecret>
token_value = <TokenValue>
token_secret = <TokenSecret>
```

## Usage

Below is the output show the attributes used.

There are two modes; Printing details about a single set `-s` or multiple sets `-f`.

Single set option will take presedence over multiple sets

```
usage: get_set.py [-h] [-s SET] [-f FILE] [-v]

optional arguments:
  -h, --help            show this help message and exit
  -s SET, --set SET
  -f FILE, --file FILE
  -v, --verbose
```

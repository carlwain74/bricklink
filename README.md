# bricklink
Utilities for collecting data from Bricklink using their API

## Requirements
- Python3.13

## Packages
- XlsWriter (Used by get_set.py)
- openpyxl (Used by get_set_openpyxl.py)
- bricklink-api (Common)
- virtualenv

## Setup Virtual Environment

### Creat virtual environment 

```
pipenv lock
pipenv sync --dev
```

## Configuration File

The file `config.ini` needs to be exist and be populated.

Head over to https://www.bricklink.com/v2/api/register_consumer.page where you can create a bricklink account and setup your API access.

Make a copy of config.ini.template
```
cp config.ini.template config.ini
```

Populate the file as follows after the the Access Token has been created for the allowedIP.
```
[secrets]
consumer_key = <ConsumerKey>
consumer_secret = <ConsumerSecret>
token_value = <TokenValue>
token_secret = <TokenSecret>
```

## Usage

There are two modes; Printing details about a single set `-s` or multiple sets `-f`.

Single set option will take presedence over multiple sets

```
usage: generate_set_sheet.py [-h] [-s SET] [-f FILE] [-v]

optional arguments:
  -h, --help            show this help message and exit
  -s SET, --set SET
  -f FILE, --file FILE
  -v, --verbose
```
### Single set

```
python generate_set_sheet.py -s 40158
2021-11-11 19:16:48 INFO     Item: 40158
2021-11-11 19:16:48 INFO       Name: Pirates Chess Set, Pirates III
2021-11-11 19:16:48 INFO       Category: Game
2021-11-11 19:16:48 INFO       Avg Price: 102 USD
2021-11-11 19:16:48 INFO       Max Price: 150 USD
2021-11-11 19:16:48 INFO       Min Price: 84 USD
2021-11-11 19:16:48 INFO       Quantity avail: 16
```

### Multiple Sets

You need to create a text file with a list of sets as follows. Script will generate a file `Items.xlsx` with a single sheet with all sets it was able to process
```
21036-1
41585-1
```
Then include the filename instead of a set.
```
pipenv run python generate_set_sheet.py -f test.txt
2021-11-11 19:18:42 INFO     Processing sets in test.txt
2021-11-11 19:18:44 INFO     Item: 21036-1
2021-11-11 19:18:44 INFO       Name: Arc De Triomphe
2021-11-11 19:18:44 INFO       Category: Architecture
2021-11-11 19:18:44 INFO       Avg Price: 91 USD
2021-11-11 19:18:44 INFO       Max Price: 99 USD
2021-11-11 19:18:44 INFO       Min Price: 75 USD
2021-11-11 19:18:44 INFO       Quantity avail: 17
2021-11-11 19:18:45 INFO     Item: 41585-1
2021-11-11 19:18:45 INFO       Name: Batman
2021-11-11 19:18:45 INFO       Category: BrickHeadz
2021-11-11 19:18:45 INFO       Avg Price: 45 USD
2021-11-11 19:18:45 INFO       Max Price: 65 USD
2021-11-11 19:18:45 INFO       Min Price: 34 USD
2021-11-11 19:18:45 INFO       Quantity avail: 13
2021-11-11 19:18:45 INFO     Total: 136USD
```

Use `get_set_openpyxl.py` to generate a sheet per set. Additional runs will add rows for each set. This is a useful aspect if you want to run the script on periodically to chart changes.

import gspread
from requests import Request, Session
import datetime
from requests.exceptions import ConnectionError, Timeout, TooManyRedirects
import json
import os

creds_path = __location__ = os.path.realpath(
    os.path.join(os.getcwd(), os.path.dirname(__file__))) + "/Crypto-Sheets.json"


gc = gspread.service_account(filename=creds_path)

# coinmarket cap credentials and data

url = 'https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest'

headers = {
    'Accepts': 'application/json',
    'X-CMC_PRO_API_KEY': '12345-abcde',
}

session = Session()
session.headers.update(headers)


def set_price_for_worksheet(worksheet, cur_bitcoin_price):
    # 1. get the name of the token
    crypto_name = worksheet.acell("B2").value

    parameters = {
        'symbol': crypto_name.upper(),
        'convert': 'EUR'
    }

    # 2. get the current trading price
    try:
        response = session.get(url, params=parameters)
        data = json.loads(response.text
                          )
        price = data['data'][crypto_name]['quote']['EUR']['price']

        # 3. set the price in the sheet
        col_values = worksheet.col_values(8)    # get the number of rows

        # Calculate the address of the last row
        coin_price_address = "H" + str((len(col_values)))
        current_date_address = "L" + str((len(col_values)))
        bitcoin_price_address = "W10"

        # set the bitcoin price value
        worksheet.update_acell(bitcoin_price_address, cur_bitcoin_price)

        # set the last price value
        worksheet.update_acell(coin_price_address, price)

        # set the current day value
        now = datetime.datetime.today()
        date_value = str(now.day) + '/' + str(now.month) + '/' + str(now.year)
        worksheet.update_acell(current_date_address, date_value)

        print("Updated values for " + str(crypto_name) + ".")
    except ConnectionError:
        print("ConnectionError from Coinbase Server")
    except Timeout:
        print("Timeout Error from Coinbase Server")
    except TooManyRedirects:
        print("TooManyRedirects from Coinbase Server")


def set_overview_sheet(overview_sheet):
    col_values = overview_sheet.col_values(11)

    # Calculate the address of the last row
    invested_address = "L" + str(len(col_values) + 1)
    profit_address = "M" + str(len(col_values) + 1)
    date_address = "K" + str(len(col_values) + 1)
    old_date_address = "K" + str(len(col_values))  # used to get the previous day

    # get the values
    old_date = overview_sheet.acell(old_date_address).value
    # remove leading zeros
    old_date_array = old_date.split("/")
    if old_date_array[0].startswith("0"):
        old_date_array[0] = old_date_array[0][1:]
    if old_date_array[1].startswith("0"):
        old_date_array[1] = old_date_array[1][1:]

    old_date = old_date_array[0] + "/" + old_date_array[1] + "/" + old_date_array[2]

    now = datetime.datetime.today()
    date_value = str(now.day) + '/' + str(now.month) + '/' + str(now.year)

    invested_amount = overview_sheet.acell("G26").value
    profit_amount = overview_sheet.acell("H26").value

    # only add a new row once a day
    if old_date != date_value:

        # set the values
        overview_sheet.update_acell(invested_address, invested_amount)
        overview_sheet.update_acell(profit_address, profit_amount)
        overview_sheet.update_acell(date_address, date_value)

        # if it is saturday, set a background color
        if datetime.datetime.today().weekday() == 5:
            overview_sheet.format(date_address + ":" + profit_address, {
                "backgroundColor": {
                    "red": 18.0,
                    "green": 153.0,
                    "blue": 226.0
                }
            })
    # if it is not a new day, only update the values
    else:
        # calculate new addresses
        invested_address = "L" + str(len(col_values))
        profit_address = "M" + str(len(col_values))

        # update the values
        overview_sheet.update_acell(invested_address, invested_amount)
        overview_sheet.update_acell(profit_address, profit_amount)

    print("Set overview sheet")


def get_bitcoin_price():
    parameters = {
        'symbol': 'BTC',
        'convert': 'EUR'
    }

    try:
        response = session.get(url, params=parameters)
        data = json.loads(response.text
                          )
        price = data['data']['BTC']['quote']['EUR']['price']

        print("Bitcoin Price: " + str(price))
        return price
    except ConnectionError:
        print("ConnectionError from Coinbase Server")
    except Timeout:
        print("Timeout Error from Coinbase Server")
    except TooManyRedirects:
        print("TooManyRedirects from Coinbase Server")


if __name__ == "__main__":
    bitcoin_price = get_bitcoin_price()
    gsheets = gc.open_by_url("https://link-to-google-spreadsheet.com")

    # don't include the first sheet
    overview_sheet = gsheets.worksheets()[0]
    worksheets = gsheets.worksheets()[1:]
    # update the values in every sheet
    for sheet in worksheets:
        set_price_for_worksheet(sheet, bitcoin_price)

    # lastly, update the overview (first sheet)
    set_overview_sheet(overview_sheet)

    print("Finished successfully")

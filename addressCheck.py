#! python3

import pandas as pd
import pgeocode
import xlwings as xw

# sheetVar: excel sheet variable; excelParam: value imported from excel; geoParam: value imported from pgeocode; columnLetter; row index
# if else statement that checks excel value against pgeocode values
def checkValue(sheetVar, excelParam, geoParam, columnLetter, row):
    if excelParam == geoParam:
        sheetVar.range(columnLetter+str(row)).value = "TRUE"
    else:
        sheetVar.range(columnLetter+str(row)).value = geoParam

# function that allows for a one at a time check
def pgeocodeCheck(countryCode,postCode):
    postal = pgeocode.Nominatim(countryCode)
    print(postal.query_postal_code(str(postCode)))

# script runs through excel sheet 
def cleanAddresses():
    data = pd.read_excel("cleanaddresses.xlsx",
                        sheet_name="Clean Addresses Test")
    wb = xw.Book('cleanaddresses.xlsx')
    addresses = wb.sheets['Clean Addresses Test']
    df = pd.DataFrame(data, columns = ['Country ISO Code','State','City','Zip/Postal Code','Address 1','Address 2'])

    for index, row in df.iterrows():
        city = row['City']
        state = row['State']
        countryCode = row['Country ISO Code']
        postCode = row['Zip/Postal Code']

        #skips US and Canada entries
        if countryCode == "US" or countryCode == "CA":
            continue
        else:
            row = index + 2

            try:
                # check pgeocode data to excel data
                postal = pgeocode.Nominatim(countryCode)
                query = postal.query_postal_code(str(postCode))
                # checks state name, at column "H"; change to desired column
                checkValue(addresses, state, query['state_name'], 'H',row)
                # checks city name, at column "J"; change to desired column
                checkValue(addresses, city, query['place_name'], 'J',row)
                # checks postal code, at column "L"; change to desired column
                checkValue(addresses, postCode, query['postal_code'], 'L',row)

            except Exception:
                print("ERROR")
                continue

def main():
    #runs clean address function
    #cleanAddresses()

    #one by one check of address info with country and postal code inputs
    pgeocodeCheck("MY",
                "40400")

main()
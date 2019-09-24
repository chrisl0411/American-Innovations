#! python3

import pandas as pd
import pgeocode
import xlwings as xw

def checkValue(sheetVar, excelParam, geoParam, columnLetter, row):
    if excelParam == geoParam:
        sheetVar.range(columnLetter+str(row)).value = "TRUE"
    else:
        sheetVar.range(columnLetter+str(row)).value = geoParam

def pgeocodeCheck(countryCode,postCode):
    postal = pgeocode.Nominatim(countryCode)
    print(postal.query_postal_code(str(postCode)))

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
                # checks state name
                checkValue(addresses, state, query['state_name'], 'H',row)
                # checks city name
                checkValue(addresses, city, query['place_name'], 'J',row)
                # checks postal code
                checkValue(addresses, postCode, query['postal_code'], 'L',row)

            except Exception:
                print("ERROR")
                continue

def main():
    cleanAddresses()
    #pgeocodeCheck("GB","NP10 8FZ")

main()
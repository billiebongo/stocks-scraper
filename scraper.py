from bs4 import BeautifulSoup
from xlrd import open_workbook
from xlutils.copy import copy

import requests
from openpyxl import load_workbook
import time

date=time.strftime("%m-%d-%y")

FILE = "HELLO_WORLD.xlsx" # 2bd: better to rename as date

INDIV_SHEETS = ["CAD", "CHF", "GBP", "JPY", "EUR", "NZD", "USD",
                "AUD", "Nikkei", "Dow Jones", "Silver", "Gold", "Oil"]

# CHICAGO MERCANTILE EXCHANGE
CURRENCIES = {
    "CAD" : 90741,
    "CHF" : 92741,
    "GBP" : 96742,
    "YEN" : 97741,
    "EURO" : 99741,
    "NZD" : 112741,
    "AUD" : 232741,
    "Nikkei" : 240741
}

# Chicago Board of Traders
TRADES = {
    "DJ": 124603
}

# CHICAGO BOARD OF COMMODITIES
COMMODITIES = {
    "SILVER" : 84691,
    "GOLD" : 88691
}

ICE ={
    "USD": 98662
}

NYME ={
    "OIL": 67651
}


def get_row_count(sheet):
    """ Find next row to append in INDIV_SHEETS. number of
    non-null values in col C is used to update the next row. """
    #some impropvement is to ensure doesnt double run
    count = 2 # cell value count starts at 2

    for cell in sheet['C']:
        if cell.value != None:
            count += 1
    return count

def update_each_sheet_NONCOMM(sheet, sheet_name, nc_long, nc_short, nc_long_week, nc_short_week, oi):
    """Update up each sheet with NON-COMMERCIAL data"""
    row_count = get_row_count(sheet)

    list1_nc = ["C", "D", "E", "F", "G", "H", "I", "J", "M", "N", "O"]
    list2_nc = ["B", "C", "D", "E", "F", "G", "H", "I", "L", "M", "N"]

    column = {"CAD": list1_nc, "CHF": list1_nc, "GBP":list2_nc, "JPY":list2_nc, "EUR":list1_nc, "NZD":list2_nc, "AUD":list2_nc,
                          "Nikkei":list2_nc, "USD":list1_nc, "Dow Jones":list2_nc, "Silver":list2_nc, "Gold":list2_nc, "Oil":list2_nc}

    # Select cells for updates, all strings
    cell1 = column[sheet_name][0]+str(row_count) #nc_long
    cell2 = column[sheet_name][1]+str(row_count) #nc_short
    cell3 = column[sheet_name][2]+str(row_count) #OI
    cell4 = column[sheet_name][3]+str(row_count) #long week
    cell5 = column[sheet_name][4]+str(row_count) #short week
    cell6 = column[sheet_name][5]+str(row_count) #real
    cell7 = column[sheet_name][6]+str(row_count) #net
    cell8 = column[sheet_name][7] + str(row_count)  # ratio
    cell9 = column[sheet_name][8] + str(row_count)  # % OI in long
    cell10 = column[sheet_name][9] + str(row_count)  # % OI in short
    cell11 = column[sheet_name][10] + str(row_count)  # %OI in net
    prev_net_cell = column[sheet_name][6]+str(row_count-1) # net previous row retrieved for value of net

    sheet[cell1] = nc_long
    sheet[cell2] = nc_short
    sheet[cell3] = oi
    sheet[cell4] = nc_long_week
    sheet[cell5] = nc_short_week

    real_pos_val =  int(nc_long_week.replace(",", "")) - int(nc_short_week.replace(",", ""))
    sheet[cell6] = real_pos_val# real position

    # Net is  previous net + current real pos
    net =  sheet[prev_net_cell].value + real_pos_val
    sheet[cell7] = net
    # Ratio is %OI long (long divide by OI)/% OI short (short divide by OI)

    # bigger over small. -ve is short > long
    # if short > long, -ve and short over long
    float_nclong = float(nc_long.replace(",", ""))
    float_ncshort = float(nc_short.replace(",", ""))


    if float_nclong >= float_ncshort:
        ratio = float_nclong / float_ncshort
    else:
        ratio = (float_ncshort/ float_nclong)*-1
    sheet[cell8] = str('%.2f' % ratio)

    sheet[cell9] = str('%.2f' % (float_nclong/float(oi.replace(",", ""))*100)) + "%"
    sheet[cell10] = str('%.2f' % (float_ncshort/float(oi.replace(",", ""))*100)) + "%"
    sheet[cell11] = str('%.2f' %( net*100/float(oi.replace(",", "")))) + "%"
    return

def update_each_sheet_COMM(sheet, sheet_name, c_long, c_short, c_long_week, c_short_week, oi):
    """Update up each individual sheet with COMMERCIAL data"""

    row_count = get_row_count(sheet) -1 # MINUS 1 because after updating non commm the row count returns the ROW
    #after the nonComm just updated


    #c_long, c_short, OI, long week, short week, real, net, ratio, % oi in long, % OI in shorts, %OI in net
    list1_c = ["S", "T", "U", "V", "W", "X", "Y", "Z", "AC", "AD", "AE"]
    list2_c = ["T", "U", "V", "W", "X", "Y", "Z", "AA", "AD", "AE", "AF"]

    column = {"CAD": list2_c, "CHF": list2_c, "GBP":list1_c, "JPY":list1_c, "EUR":list2_c, "NZD":list1_c, "AUD":list1_c,
                          "Nikkei":list1_c, "USD":list2_c, "Dow Jones":list1_c, "Silver":list1_c, "Gold":list1_c, "Oil":list1_c}


    # Select cells for updates, all strings
    cell1 = column[sheet_name][0]+str(row_count) #c_long
    cell2 = column[sheet_name][1]+str(row_count) #c_short
    cell3 = column[sheet_name][2]+str(row_count) #OI
    cell4 = column[sheet_name][3]+str(row_count) #long week
    cell5 = column[sheet_name][4]+str(row_count) #short week
    cell6 = column[sheet_name][5]+str(row_count) #real
    cell7 = column[sheet_name][6]+str(row_count) #net
    cell8 = column[sheet_name][7] + str(row_count)  # ratio
    cell9 = column[sheet_name][8] + str(row_count)  # % OI in long
    cell10 = column[sheet_name][9] + str(row_count)  # % OI in short
    cell11 = column[sheet_name][10] + str(row_count)  # %OI in net
    prev_net_cell = column[sheet_name][6]+str(row_count-1) # net previous row retrieved for value of net

    sheet[cell1] = c_long
    sheet[cell2] = c_short
    sheet[cell3] = oi

    sheet[cell4] = c_long_week
    sheet[cell5] = c_short_week

    real_pos_val =  int(c_long_week.replace(",", "")) - int(c_short_week.replace(",", ""))
    sheet[cell6] = real_pos_val# real position

    # Net is  previous net + current real pos
    net =  sheet[prev_net_cell].value + real_pos_val
    sheet[cell7] = net
    # Ratio is %OI long (long divide by OI)/% OI short (short divide by OI)

    # bigger over small. -ve is short > long
    # if short > long, -ve and short over long
    float_clong = float(c_long.replace(",", ""))
    float_cshort = float(c_short.replace(",", ""))

    if float_clong >= float_cshort:
        ratio = float_clong / float_cshort
    else:
        ratio = (float_cshort/ float_clong)*-1
    sheet[cell8] = '%.2f' % ratio
    sheet[cell9] = str('%.2f' % (float_clong/float(oi.replace(",", ""))*100)) + "%"
    sheet[cell10] = str('%.2f' % (float_cshort/float(oi.replace(",", ""))*100)) + "%"

    sheet[cell11] = str('%.2f' % (net*100/float(oi.replace(",", "")))) + "%"

    return

def update_dates(wb,date_string): #manually pass as arg
    """ Update date cells for all sheets """
    columns = {"CAD": ["B", "S"], "CHF": ["B", "S"], "GBP":["A", "R"], "JPY":["A", "R"], "EUR":["B", "S"], "NZD":["A", "R"], "AUD":["A", "R"],
                          "Nikkei":["A", "R"], "USD":["B", "S"], "Dow Jones":["A", "R"], "Silver":["A", "R"], "Gold":["A", "R"], "Oil":["A", "R"]
    }

    wb['Main']['A1'] = date_string
    for sheet_name, col in columns.items():
        cell_nc = col[0] + str(get_row_count(wb[sheet_name])-1)
        wb[sheet_name][cell_nc] =  date_string
        cell_c = col[1] + str(get_row_count(wb[sheet_name]) - 1)
        print(cell_c)
        wb[sheet_name][cell_c] = date_string


    return


def update_all_sheets(wb ,curr_dict):
    """ Calls functions to update each individual sheet with NON-COMM and COMM data """
    for sheet_name in INDIV_SHEETS:

        dict_value = {"CAD":"CAD", "CHF":"CHF", "GBP":"GBP", "JPY":"YEN", "EUR":"EURO", "NZD":"NZD", "AUD":"AUD",
                          "Nikkei":"Nikkei", "USD":"USD", "Dow Jones":"DJ", "Silver":"SILVER", "Gold":"GOLD", "Oil":"OIL" }# matches sheet_name to dict_name
        print("sheet_name: {}, dict_name: {}".format(sheet_name, dict_value[sheet_name]))
        params=curr_dict[dict_value[sheet_name]]

        #sheet, sheet_name, nc_long, nc_short, nc_long_week, nc_short_week, oi)
        #  ### CURR DICT KEY MIGHT NOT BESAME AS SHEETNAME
        # params => [nc_long, nc_short, c_long, c_short, oi, nc_long_week, nc_short_week ]

        update_each_sheet_NONCOMM(wb[sheet_name], sheet_name,params[0],params[1], params[5], params[6], params[4])
        #(sheet, sheet_name, c_long, c_short, c_long_week, c_short_week, oi

        print(wb[sheet_name], sheet_name,params[2],params[3], params[5], params[6], params[4])
        update_each_sheet_COMM(wb[sheet_name], sheet_name,params[2],params[3], params[7], params[8], params[4])


    return


def get_dets(CAD, s): #2bd: CAD is a bad var name
    """ Parse HTML to retrieve data """
    try:
        row_all = s.split(str(CAD))[1].split('Changes')[0].split('All')[1].split('Old')[0].replace(":", "")
        integers = row_all.split()
        oi = integers[0]
        nc_long = integers[1]
        nc_short = integers[2]
        c_long = integers[4]
        c_short = integers[5]
        this_week_row = s.split(str(CAD))[1].split('Changes')[1].split('Percent')[0].split(':')[4]
        integers2 = this_week_row.split()


        nc_long_week =integers2[0]
        nc_short_week =integers2[1]


        if len(integers2)== 7:
            c_long_week = integers2[3]
            c_short_week =integers2[4]
        else:
            print("There arent 7 items in the changes line, format has changed!")
            raise


        return nc_long, nc_short, c_long, c_short, oi, nc_long_week, nc_short_week, c_long_week, c_short_week
    except Exception as e:
        print(CAD)
        print("ERROR")
        print(e)


        raise


def get_html(link):
    """ GET request to URL and returns text in HTML """

    r = requests.get(link)
    soup = BeautifulSoup(r.content)
    data=soup.findAll('pre')
    return data[0].text



def main_sheet():
    """
    Retrieves data from respective URLs and returns curr_dict which contains
    nc_long, nc_short, c_long, c_short, oi, nc_long_week, nc_short_week, c_long_week, c_short_week
    for each item in CURRENCIES, TRADES, COMMODOTIES, ICE, NYME
    """

    data_string = get_html('http://www.cftc.gov/dea/futures/deacmelf.htm')
    curr_dict = {}
    for c, val in CURRENCIES.items():

        nc_long, nc_short, c_long, c_short, oi, nc_long_week, nc_short_week, c_long_week, c_short_week = get_dets(val, data_string)
        curr_dict[c] = [nc_long, nc_short, c_long, c_short, oi, nc_long_week, nc_short_week, c_long_week, c_short_week ]

    data_string = get_html('http://www.cftc.gov/dea/futures/deacbtlf.htm')
    for c, val in TRADES.items():

        nc_long, nc_short, c_long, c_short, oi, nc_long_week, nc_short_week, c_long_week, c_short_week  = get_dets(val, data_string)
        curr_dict[c] = [ nc_long, nc_short, c_long, c_short, oi, nc_long_week, nc_short_week, c_long_week, c_short_week ]

    data_string = get_html('http://www.cftc.gov/dea/futures/deacmxlf.htm')
    for c, val in COMMODITIES.items():
        nc_long, nc_short, c_long, c_short, oi, nc_long_week, nc_short_week, c_long_week, c_short_week  = get_dets(val, data_string)
        curr_dict[c] = [nc_long, nc_short, c_long, c_short,oi, nc_long_week, nc_short_week, c_long_week, c_short_week ]

    data_string = get_html('http://www.cftc.gov/dea/futures/deanybtlf.htm')
    for c, val in ICE.items():
        nc_long, nc_short, c_long, c_short, oi, nc_long_week, nc_short_week, c_long_week, c_short_week  = get_dets(val, data_string)
        curr_dict[c] = [nc_long, nc_short, c_long, c_short,oi, nc_long_week, nc_short_week, c_long_week, c_short_week ]

    data_string = get_html('http://www.cftc.gov/dea/futures/deanymelf.htm')
    for c, val in NYME.items():
        nc_long, nc_short, c_long, c_short, oi, nc_long_week, nc_short_week, c_long_week, c_short_week  = get_dets(val, data_string)
        curr_dict[c] = [nc_long, nc_short, c_long, c_short, oi, nc_long_week, nc_short_week, c_long_week, c_short_week ]
    print(curr_dict)
    return curr_dict



def insert_excel(curr_dict, date_string):
    """
    Insert retrieved data in ALL SHEETS
    Saves File
    """

    wb = load_workbook(filename=FILE)
    Main = wb['Main']

    #update main sheet
    #CANADA
    Main['D3'], Main['E3'], Main['H3'], Main['I3'], Main['L3'] = curr_dict["CAD"][0], curr_dict["CAD"][1], curr_dict["CAD"][2], curr_dict["CAD"][3], curr_dict["CAD"][4]
    #update_each_sheet( wb['CHF'], curr_dict["CAD"][0], curr_dict["CAD"][1], curr_dict["CAD"][2], curr_dict["CAD"][3], curr_dict["CAD"][4])
    #CHF

    Main['D4'], Main['E4'], Main['H4'], Main['I4'], Main['L4'] = curr_dict["CHF"][0], curr_dict["CHF"][1], curr_dict["CHF"][2], curr_dict["CHF"][3], curr_dict["CHF"][4]
    #GBP
    Main['D5'], Main['E5'], Main['H5'], Main['I5'], Main['L5'] = curr_dict["GBP"][0], curr_dict["GBP"][1], curr_dict["GBP"][2], curr_dict["GBP"][3], curr_dict["GBP"][4]
    #YEN
    Main['D6'], Main['E6'], Main['H6'], Main['I6'], Main['L6'] = curr_dict["YEN"][0], curr_dict["YEN"][1], curr_dict["YEN"][2], curr_dict["YEN"][3], curr_dict["YEN"][4]
    #EURO
    Main['D7'], Main['E7'], Main['H7'], Main['I7'], Main['L7'] = curr_dict["EURO"][0], curr_dict["EURO"][1], curr_dict["EURO"][2], curr_dict["EURO"][3], curr_dict["EURO"][4]
    #NZD
    Main['D8'], Main['E8'], Main['H8'], Main['I8'], Main['L8'] = curr_dict["NZD"][0], curr_dict["NZD"][1], curr_dict["NZD"][2], curr_dict["NZD"][3], curr_dict["NZD"][4]
    #AUD
    Main['D9'], Main['E9'], Main['H9'], Main['I9'], Main['L9'] = curr_dict["AUD"][0], curr_dict["AUD"][1], curr_dict["AUD"][2], curr_dict["AUD"][3], curr_dict["AUD"][4]
    #Nikkei
    Main['D10'], Main['E10'], Main['H10'], Main['I10'], Main['L10'] = curr_dict["Nikkei"][0], curr_dict["Nikkei"][1], curr_dict["Nikkei"][2], curr_dict["Nikkei"][3], curr_dict["Nikkei"][4]
    #DJ
    Main['D11'], Main['E11'], Main['H11'], Main['I11'], Main['L11'] = curr_dict["DJ"][0], curr_dict["DJ"][1], curr_dict["DJ"][2], curr_dict["DJ"][3], curr_dict["DJ"][4]
    #SILVER
    Main['D12'], Main['E12'], Main['H12'], Main['I12'], Main['L12'] = curr_dict["SILVER"][0], curr_dict["SILVER"][1], curr_dict["SILVER"][2], curr_dict["SILVER"][3], curr_dict["SILVER"][4]
    #GOLD
    Main['D13'], Main['E13'], Main['H13'], Main['I13'], Main['L13'] = curr_dict["GOLD"][0], curr_dict["GOLD"][1], curr_dict["GOLD"][2], curr_dict["GOLD"][3], curr_dict["GOLD"][4]
    #OIL
    Main['D14'], Main['E14'], Main['H14'], Main['I14'], Main['L14'] = curr_dict["OIL"][0], curr_dict["OIL"][1], curr_dict["OIL"][2], curr_dict["OIL"][3], curr_dict["OIL"][4]

    #USD
    Main['D15'], Main['E15'], Main['H15'], Main['I15'], Main['L15'] = curr_dict["USD"][0], curr_dict["USD"][1], curr_dict["USD"][2], curr_dict["USD"][3], curr_dict["USD"][4]

    # Update the other sheets
    update_all_sheets(wb, curr_dict)


    update_dates(wb, date)

    #only save when error free
    wb.save(FILE)

    return

if __name__ == "__main__":
    curr_dict=main_sheet()
    insert_excel(curr_dict, date)





import requests
import json
import csv
import os
# from datetime import datetime
from datetime import date
import urllib.request
import shutil
from config import *
import glob

# month to select from when running
monthTotal = {
    'January': 1,
    'February': 2,
    'March': 3,
    'April': 4,
    'Mai': 5,
    'June': 6,
    'July': 7,
    'August': 8,
    'September': 9,
    'October': 10,
    'November': 11,
    'Dezember': 12
}

# header of csv-file
header = {
    'Umsatz (ohne Soll/Haben-Kz)': '',
    'Soll/Haben-Kennzeichen': '',
    'WKZ Umsatz': '',
    'Kurs': '',
    'Basis-Umsatz': '',
    'WKZ Basis-Umsatz': '',
    'Konto': '',
    'Gegenkonto (ohne BU-Schlüssel)': '',
    'BU-Schlüssel': '',
    'Belegdatum': '',
    'Belegfeld 1': '',
    'Belegfeld 2': '',
    'Skonto': '',
    'Buchungstext': '',
    'Postensperre': '',
    'Diverse Adressnummer': '',
    'Geschäftspartnerbank': '',
    'Sachverhalt': '',
    'Zinssperre': '',
    'Beleglink': '',
    'Beleginfo - Art 1': '',
    'Beleginfo - Inhalt 1': '',
    'Beleginfo - Art 2': '',
    'Beleginfo - Inhalt 2': '',
    'Beleginfo - Art 3': '',
    'Beleginfo - Inhalt 3': '',
    'Beleginfo - Art 4': '',
    'Beleginfo - Inhalt 4': '',
    'Beleginfo - Art 5': '',
    'Beleginfo - Inhalt 5': '',
    'Beleginfo - Art 6': '',
    'Beleginfo - Inhalt 6': '',
    'Beleginfo - Art 7': '',
    'Beleginfo - Inhalt 7': '',
    'Beleginfo - Art 8': '',
    'Beleginfo - Inhalt 8': '',
    'KOST1 - Kostenstelle': '',
    'KOST2 - Kostenstelle': '',
    'Kost-Menge': '',
    'EU-Land u. UStID': '',
    'EU-Steuersatz': '',
    'Abw. Versteuerungsart': '',
    'Sachverhalt L+L': '',
    'Funktionsergänzung L+L': '',
    'BU 49 Hauptfunktionstyp': '',
    'BU 49 Hauptfunktionsnummer': '',
    'BU 49 Funktionsergänzung': '',
    'Zusatzinformation - Art 1': '',
    'Zusatzinformation- Inhalt 1': '',
    'Zusatzinformation - Art 2': '',
    'Zusatzinformation- Inhalt 2': '',
    'Zusatzinformation - Art 3': '',
    'Zusatzinformation- Inhalt 3': '',
    'Zusatzinformation - Art 4': '',
    'Zusatzinformation- Inhalt 4': '',
    'Zusatzinformation - Art 5': '',
    'Zusatzinformation- Inhalt 5': '',
    'Zusatzinformation - Art 6': '',
    'Zusatzinformation- Inhalt 6': '',
    'Zusatzinformation - Art 7': '',
    'Zusatzinformation- Inhalt 7': '',
    'Zusatzinformation - Art 8': '',
    'Zusatzinformation- Inhalt 8': '',
    'Zusatzinformation - Art 9': '',
    'Zusatzinformation- Inhalt 9': '',
    'Zusatzinformation - Art 10': '',
    'Zusatzinformation- Inhalt 10': '',
    'Zusatzinformation - Art 11': '',
    'Zusatzinformation- Inhalt 11': '',
    'Zusatzinformation - Art 12': '',
    'Zusatzinformation- Inhalt 12': '',
    'Zusatzinformation - Art 13': '',
    'Zusatzinformation- Inhalt 13': '',
    'Zusatzinformation - Art 14': '',
    'Zusatzinformation- Inhalt 14': '',
    'Zusatzinformation - Art 15': '',
    'Zusatzinformation- Inhalt 15': '',
    'Zusatzinformation - Art 16': '',
    'Zusatzinformation- Inhalt 16': '',
    'Zusatzinformation - Art 17': '',
    'Zusatzinformation- Inhalt 17': '',
    'Zusatzinformation - Art 18': '',
    'Zusatzinformation- Inhalt 18': '',
    'Zusatzinformation - Art 19': '',
    'Zusatzinformation- Inhalt 19': '',
    'Zusatzinformation - Art 20': '',
    'Zusatzinformation- Inhalt 20': '',
    'Stück': '',
    'Gewicht': '',
    'Zahlweise': '',
    'Forderungsart': '',
    'Veranlagungsjahr': '',
    'Zugeordnete Fälligkeit': '',
    'Skontotyp': '',
    'Auftragsnummer': '',
    'Buchungstyp': '',
    'USt-Schlüssel (Anzahlungen)': '',
    'EU-Land (Anzahlungen)': '',
    'Sachverhalt L+L (Anzahlungen)': '',
    'EU-Steuersatz (Anzahlungen)': '',
    'Erlöskonto (Anzahlungen)': '',
    'Herkunft-Kz': '',
    'Buchungs GUID': '',
    'KOST-Datum': '',
    'SEPA-Mandatsreferenz': '',
    'Skontosperre': '',
    'Gesellschaftername': '',
    'Beteiligtennummer': '',
    'Identifikationsnummer': '',
    'Zeichnernummer': '',
    'Postensperre bis': '',
    'Bezeichnung SoBil-Sachverhalt': '',
    'Kennzeichen SoBil-Buchung': '',
    'Festschreibung': '',
    'Leistungsdatum': '',
    'Datum Zuord. Steuerperiod': ''
}

accounts = {
    'collective': '10000',
    '19': '8400',
    '7': '8300'
}

i = 0
offset = 0

# orders in 100 steps / set 15 here to download the latest 1500 invoices
countJSONData = 100

countZero = 0
countNotPaid = 0
countTrueOrder = 0
countCanceled = 0

# csv writer function
def csv_writer(month, header):
    with open("HoertHinGmbH_EcwidOrders" + str(month) + ".csv", "a", encoding='UTF8', newline='') as csvfile:
        writer = csv.writer(csvfile, delimiter= ';')
        # writer = csv.writer(csvfile, delimiter= ',')
        writer.writerow(header.values())
    csvfile.close()

# print("\n")
# # waiting for user input - choose month
# for months,number in monthTotal.items():
#     print(str(number).zfill(2), months)

today = date.today()
month = int(str(today).split()[0].split('-')[1]) - 1
if month == 0:
    year = int(str(today).split('-')[0]) - 1
    month = 12

print("getting ecwid data for month:", month, "on",today)
# while True:
#     month = int(input("\n>>> Select a month: "))
#     if month > today.month:
#         print("selected month has to be in the past. try again!")
#         continue
#     else:
#         break

# make folder for invoices
folder = "Ecwid_Rechnungen_" + str(month) + "_" + str(date.today())
if not os.path.exists(folder):
    os.mkdir(folder)
    print(folder)
else:
    print(folder, "exists!")

# make canceled for invoices
canceledFolder = "Storno_" + str(month) + "_" + str(date.today())
if not os.path.exists(canceledFolder):
    os.mkdir(canceledFolder)
    print(canceledFolder)
else:
    print(canceledFolder, "exists!")

# read json-files
k = 0
with open("HoertHinGmbH_EcwidOrders" + str(month) + ".csv", "a", encoding='UTF8', newline='') as csvfile:
    writer = csv.writer(csvfile, delimiter= ';')
    # writer = csv.writer(csvfile, delimiter= ',')
    writer.writerow(header.keys())
csvfile.close()

# main loop
# get data from ecwid api and store to json
while True:
    response = requests.get("https://app.ecwid.com/api/v3/"+str(id)+"/orders?offset="+str(offset)+"&limit=100&token="+str(token))
    data = response.json()
    json_object = json.dumps(data, indent=4, ensure_ascii=False)
    with open("data"+str(i)+".json", "w") as outfile:
        outfile.write(json_object)
    # print(data)
    outfile.close()

    while True:
        # open json for reading only
        with open("data"+str(i)+".json", "r") as infile:
            print("data"+str(i)+".json")
            data = json.loads(infile.read())
            count = data['count']
            v = 0
            for v in range(count):
                try:
                    # get time variables
                    orderYear = int(data['items'][v]['invoices'][0]['created'].split('-')[0])
                    orderMonth = int(data['items'][v]['invoices'][0]['created'].split('-')[1])
                    orderDate = int(data['items'][v]['invoices'][0]['created'].split('-')[2].split(' ')[0])
                    fullDate = str(orderDate).zfill(2) + "." + str(orderMonth).zfill(2) + "." + str(orderYear)
                except:
                    pass

                # set counters for printing in prompt only
                if orderMonth == month:
                    if data['items'][v]['paymentStatus'] != 'PAID':
                        countNotPaid += 1
                    if data['items'][v]['paymentStatus'] == 'CANCELLED':
                        countCanceled += 1

                    if data['items'][v]['total'] == 0:
                        countZero += 1
                        # print("invoice to skip:", data['items'][v]['id'], (str(data['items'][v]['total']) + "€ <<<"), data['items'][v]['billingPerson']['name'])
                        pass

                    else:
                        # get values from json and store into dictionary
                        # # format 4.0 -> 4,0
                        total_f = str(format(data['items'][v]['total'], '.2f')).replace('.', ',')
                        header['Umsatz (ohne Soll/Haben-Kz)'] = total_f
                        header['Soll/Haben-Kennzeichen'] = "H"
                        header['Gegenkonto (ohne BU-Schlüssel)'] = accounts['collective']
                        header['Belegdatum'] = fullDate
                        header['Belegfeld 1'] = data['items'][v]['invoices'][0]['id']
                        # header['Festschreibung'] = 0
                        # check how many purchases in invoice
                        res = len([ele for ele in data['items'][v]['items'] if isinstance(ele, dict)])
                        m = 0
                        # loop for more then one purchase per invoice, using 'm' as loop-count
                        if res > 1:
                            for m in range(res):
                                # ignore 0€ invoices
                                if data['items'][v]['items'][m]['price'] == 0:
                                    pass

                                else:
                                    # coupon activated?
                                    try:
                                        if data['items'][v]['items'][m]['couponAmount'] == True:
                                            if data['items'][v]['items'][m]['price'] == data['items'][v]['items'][m]['couponAmount']:
                                                pass

                                    except:
                                        # if more then one purchase per invoice
                                        try:
                                            # if 19%
                                            if data['items'][v]['items'][m]['taxes'][0]['value'] == 19:
                                                # format 4.0 -> 4,0
                                                total_f = str(format(data['items'][v]['items'][m]['price'], '.2f')).replace('.', ',')
                                                header['Umsatz (ohne Soll/Haben-Kz)'] = total_f
                                                header['Konto'] = accounts['19']
                                                # append to csv
                                                csv_writer(month, header)
                                            # if 7%
                                            elif data['items'][v]['items'][m]['taxes'][0]['value'] == 7:
                                                total_f = str(format(data['items'][v]['items'][m]['price'], '.2f')).replace('.', ',')
                                                header['Umsatz (ohne Soll/Haben-Kz)'] = total_f
                                                header['Konto'] = accounts['7']
                                                # append to csv
                                                csv_writer(month, header)

                                        except:
                                            # error
                                            header['Konto'] = "ERROR"
                                            # append to csv
                                            csv_writer(month, header)
                                m += 1
                        else:
                            # if only one purchase per invoice
                            try:
                                taxes = format(data['items'][v]['total'] / data['items'][v]['totalWithoutTax'], '.2f')
                                if taxes == "1.19":
                                    header['Konto'] = accounts['19']
                                elif taxes == "1.07":
                                    header['Konto'] = accounts['7'] 

                            except:
                                # error
                                print("error - division by zero not possible")
                                header['Konto'] = None
                                header['Gegenkonto (ohne BU-Schlüssel)'] = None

                            # append to csv, when only one product in invoice
                            csv_writer(month, header)

                        # # print status in terminal only
                        # print(">>> invoice ready for download:", str(data['items'][v]['invoices'][0]['id']), "orders:", res)
                        countTrueOrder += 1

                    v += 1
                    if v == count:
                        break
        k += 1
        infile.close()
        break
    i += 1
    offset += 100

    if orderMonth < month:
        break

print("\n-----------------------------------------------------------------")
print("bills not paid:", countNotPaid)
print("0€ skipped invoices:", countZero)
print("canceld counter:", countCanceled)
print("orders transfered to csv:", countTrueOrder, "\n")
print("csv-file ready for upload!")
print("-----------------------------------------------------------------\n")

# pdf download
jsonCounter = len(glob.glob1("../SansSibar","*.json"))
# print(jsonCounter)

i = 0
countOrders = 1

for j in range(jsonCounter):
    with open("data"+str(i)+".json", "r") as infile:
        data = json.loads(infile.read())
        count = data['count']
        for v in range(count):
            try:
                # get time variables
                orderMonth = int(data['items'][v]['invoices'][0]['created'].split('-')[1])
            except:
                pass

            # set counters for printing in prompt only
            if orderMonth == month:
                if data['items'][v]['total'] == 0:
                    countZero += 1
                    # print("invoice skipped:", data['items'][v]['id'], (str(data['items'][v]['total']) + "€ <<<"), data['items'][v]['billingPerson']['name'])
                    pass
                else:
                    # download pdf files
                    pdf_file = "Rechnung_" + str(data['items'][v]['invoices'][0]['id']) + "_für_Bestellung_" + str(data['items'][v]['id'])
                    pdf_path = data['items'][v]['invoices'][0]['link']

                    def download_file(download_url, filename):
                        response = urllib.request.urlopen(download_url)    
                        file = open(pdf_file + ".pdf", 'wb')
                        file.write(response.read())
                        file.close()

                    # move pdf file to new folder
                    download_file(pdf_path, " ")
                    print((str(countOrders)+"/"+str(countTrueOrder)), "downloading pdf:", pdf_file)
                    countOrders +=1 
                    shutil.move((pdf_file + ".pdf"), (folder + "/" + pdf_file + ".pdf"))

                    if data['items'][v]['paymentStatus'] == 'CANCELLED':
                        canceled = len(data['items'][v]['invoices'])
                        for k in range(canceled):
                            pdf_path = data['items'][v]['invoices'][k]['link']
                            if data['items'][v]['invoices'][k]['type'] == 'FULL_CANCEL':
                                pdf_file = "Storno_" + str(data['items'][v]['invoices'][k]['id']) + "_für_Bestellung_" + str(data['items'][v]['id'])
                                print((str(countOrders)+"/"+str(countTrueOrder)), "downloading pdf:", pdf_file)
                                download_file(pdf_path, "Test")
                                shutil.move((pdf_file + ".pdf"), (canceledFolder + "/" + pdf_file + ".pdf"))
                            k += 1
    i += 1
    if orderMonth == month -2:
        break

# print status in terminal only
shutil.move(("HoertHinGmbH_EcwidOrders" + str(month) + ".csv"), (folder + "/HoertHinGmbH_EcwidOrders" + str(month) + ".csv"))
print("\n-----------------------------------------------------------------")
print("csv-file has been moved to folder " + "/" + folder)
print("\njob done!" +fullDate+ "\n")

# remove json-files
i = 0
while True:
    if os.path.exists("data"+str(i)+".json"):
        os.remove("data"+str(i)+".json")
    else:
        # print("all json-files removed!")
        # print("-----------------------------------------------------------------\n")
        break
    i += 1

exit(0)
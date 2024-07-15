import pandas as pd
from datetime import datetime
import airporttime
import streamlit as st
from ics import Calendar, Event

file = pd.read_excel('00238975_7629813488369943441.xls',
                     sheet_name='RosterReport',
                     usecols="B,G,J,L,M,N,O,S,U,V,X, AA,AB,AD,AE,AF",
                     skiprows=9)
header = file.columns.str.replace('\n', ' ')
file.columns = header

data = file.dropna(how='all').copy()
data['from'] = data['Sector'].str.replace(r'-.*', '', regex=True)
data['to'] = data['Sector'].str.replace(r'.*-', '', regex=True)


def getUTC(dt, place):
    if not isinstance(dt, datetime):
        return
    apt = airporttime.AirportTime(iata_code=place)
    return apt.to_utc(dt)


print(getUTC(datetime.now(), "SIN"))
print(getUTC("", "SIN"))

data_utc = data.copy()
data_utc["RptUTC"] = data.apply(lambda x: getUTC(x["Rpt"], x["from"]), axis=1)
data_utc["STDUTC"] = data.apply(lambda x: getUTC(x["STD"], x["from"]), axis=1)
data_utc["STAUTC"] = data.apply(lambda x: getUTC(x["STA"], x["to"]), axis=1)

c = Calendar()
e = None

for index, row in data_utc.iterrows():
    if not pd.isnull(row['STDUTC']):
        flt_num = "" if pd.isnull(row['Flight Number']) else row['Flight Number'] + " "
        e = Event()
        e.name = f"{row['Duty']} {flt_num}{row['Sector']}"
        e.begin = row['STDUTC']
    if not pd.isnull(row['STAUTC']):
        e.end = row['STAUTC']
        c.events.add(e)
        e = None
print(c.serialize())
with open('sortie.ics', 'w') as my_file:
    my_file.writelines(c.serialize_iter())

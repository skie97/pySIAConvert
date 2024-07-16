import pandas as pd
from datetime import datetime
import airporttime
from ics import Calendar, Event


def getUTC(dt, place):
    if not isinstance(dt, datetime):
        return
    apt = airporttime.AirportTime(iata_code=place)
    return apt.to_utc(dt)


def convert_file(input):
    file = pd.read_excel(input,
                         sheet_name='RosterReport',
                         usecols="B,G,J,L,M,N,O,S,U,V,X, AA,AB,AD,AE,AF",
                         skiprows=9)
    header = file.columns.str.replace('\n', ' ')
    file.columns = header

    data = file.dropna(how='all').copy()
    data['from'] = data['Sector'].str.replace(r'-.*', '', regex=True)
    data['to'] = data['Sector'].str.replace(r'.*-', '', regex=True)

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
    return c.serialize()
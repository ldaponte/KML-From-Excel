# This program is used to read addresses from an Excel file, 
# look up those addresses using Google Maps API, and generate a KML file
# If the Opt-In column contains a 'Y', then a precise point is mapped
# otherwise the lat and long are truncated to reduce precision
# See code below for requierd Excel columns

from polycircles import polycircles
import simplekml
import requests
import json
import pandas as pd
import re
import os

def truncate(f, n):
    '''Truncates/pads a float f to n decimal places without rounding'''
    s = '{}'.format(f)
    if 'e' in s or 'E' in s:
        return '{0:.{1}f}'.format(f, n)
    i, p, d = s.partition('.')
    return '.'.join([i, (d+'0'*n)[:n]])

def safeget(dct, *keys):
    for key in keys:
        try:
            dct = dct[key]
        except KeyError:
            return None
    return dct

url = 'https://maps.googleapis.com/maps/api/geocode/json'
key = os.getenv('GOOGLE_MAPS_API_KEY') # Make sure to set this environment variable to the Google Maps API key

roster = pd.read_excel('addresses.xlsx', dtype=str, na_filter = False)
roster.fillna('')

kml_optin = simplekml.Kml()
kml_optout = simplekml.Kml()

for index, row in roster.iterrows():

    row.fillna('')
    address = str(row['Address'] + ', ' + row['City '] + ', ' + row['State'] + ' ' + row['ZipCode']).upper()
    address = re.sub(' +', ' ', address).strip()
    callsign = row['Call']
    optin = row['Opt-In']
    print(address)

    query = {'parameters':'','address':address, 'key':key}
    response = requests.get(url, params=query)

    data = json.loads(response.content)

    if data['status'] == 'OK':

        lat = safeget(data, 'results', 0, 'geometry', 'location', 'lat')
        long = safeget(data, 'results', 0, 'geometry', 'location', 'lng')

        if optin == 'Y':

            f_lat = float(lat)
            f_long = float(long)

            pnt = kml_optin.newpoint()
            pnt.name = callsign
            pnt.coords = [(f_long, f_lat)]
            pnt.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/pal4/icon57.png'

        else:
            f_lat = float(truncate(float(lat), 2))
            f_long = float(truncate(float(long), 2))
            radius = 800
            color = simplekml.Color.orange

            # Note: radious is in meters
            polycircle = polycircles.Polycircle(latitude=f_lat,
                                                longitude=f_long,
                                                radius=radius,
                                                number_of_vertices=36)

            pol = kml_optout.newpolygon(name=callsign, outerboundaryis=polycircle.to_kml())
            pol.style.polystyle.color = simplekml.Color.changealphaint(200, color)

            pnt = kml_optout.newpoint()
            pnt.name = callsign
            pnt.coords = [(f_long, f_lat)]
            pnt.style.iconstyle.icon = simplekml.Icon(None)

    else:
        print('Maps API status: ' + data['status'])

opt_in_file = './opt_in.kml'
opt_out_file = './opt_out.kml'

print('Saving: ' + opt_in_file)
kml_optin.save(opt_in_file)

print('Saving: ' + opt_out_file)
kml_optout.save(opt_out_file)
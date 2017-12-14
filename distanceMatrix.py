#!/usr/bin/env python

import pandas as pd
import numpy as np
import googlemaps
from tqdm import tqdm

def save_matrix():
    df.to_excel(writer, sheet_name='Matrix', index=False)
    writer.save()
    writer.close()

#config
filename = '<< PATH TO EXCEL FILE WITH STARTING CITIES >>' #input data
writer = pd.ExcelWriter('<< PATH TO RESULT FILE >>',
                        engine='xlsxwriter') # output file
gmaps = googlemaps.Client(key='<< YOUR API KEY>>') # paste your Google maps API key here
to_cities = ('Bydgoszcz', 'Chorzów', 'Gdańsk',
             'Gdynia', 'Opole', 'Poznań',
             'Szczecin', 'Toruń', 'Wrocław') # destinations

#Program start
try:
    df = pd.read_excel(filename, header=0) # load to DataFrame
except:
    print('Error opening: {}'.format(filename))
print('Opening: {}'.format(filename))

for city in to_cities:
    df[city] = np.nan # add destination columns to dataframe

for row in tqdm(df.itertuples()):
    start = row.<< DATAFRAME COLUMN NAME FOR SOURCE CITY >>
    poz = row.Index
    for city in to_cities:
        if city == start :
            dist_km = 0
            df.loc[poz, city] = dist_km
        else:
            google = gmaps.distance_matrix(start, city, units='metric') # you can change units here
            dist_m = google['rows'][0]['elements'][0]['distance']['value'] # load meters value
            dist_km = round(dist_m / 1000, 2) # change to kilometers
            df.loc[poz, city] = dist_km

save_matrix()

print('All done!')

#!/usr/bin/env python3
import pandas as pd
import numpy as np
import csv

lager = pd.read_excel('GEWERBE - Rechnung - Lager.xlsx',
                      header=2,
                      dtype={
                          "Jahr" : str,
                      })

# remove the stuff below in the Lager file
lager = lager.iloc[0:lager['Jahr'].last_valid_index()+1]
lager = lager.convert_dtypes({
    'verkauft' : np.float64,
})

# Load Lager and remove all rows where Artikelnr is not set.
# Ignore rows that have an invalid value in the Artikelnr column
lager_agg = lager[lager['eingestellt'].isnull()].groupby('Artikelnr', dropna=True).agg(
    available=('Stk', 'sum'),
    price=('verkauft', 'max'))
lager[lager['Artikelnr'] == 'DEU2015002320']['verkauft'].max()

# Load Coin Database
coin_db = pd.read_excel('Münzdatenbank.xlsx',
                        dtype={
                            'groupid': str,
                            'Praegejahr': str
                            })

# Generate the Picpath
coin_db['picpath'] = coin_db['Artikelnr'] + '0.jpg'

# Append the required 0 and 2 to the end of the Artikelnr
coin_db_2 = coin_db.copy()
coin_db_2['Artikelnr'] = coin_db['Artikelnr'] + '2'
coin_db_2['rate'] = "20"

coin_db['Artikelnr'] = coin_db['Artikelnr'] + '0'
coin_db['rate'] = "0"

# join the dataframes to get descriptions from coin_db
coin_db = pd.concat([coin_db, coin_db_2])
coin_db = coin_db.set_index('Artikelnr')
coins_in_shop = lager_agg.join(coin_db, how="right")
coins_in_shop.loc[coins_in_shop['available'].notna(), 'CMD'] = ''
coins_in_shop['available'] = coins_in_shop['available'].fillna(0)

# get missing values on the right
coin_db_missing = lager_agg.join(coin_db, how="left")
coin_db_missing = coin_db_missing.loc[coin_db_missing['name'].isna(), ['available', 'price']
                                      ].join(lager.set_index('Artikelnr')['Münzbezeichnung'],
                                             how='left').reset_index().groupby('Münzbezeichnung').first()
coin_db_missing.to_excel("fehlt.xlsx")

print("Fehlende Münzen in der Münzdatenbank:")
print(coin_db_missing)

coins_in_shop.to_csv('gesamt.csv',
                     sep='\t',
                     encoding="cp1252",
                     line_terminator='\r\n',
                     quoting=csv.QUOTE_MINIMAL,
                     quotechar='\'',
                     index=True,
                     index_label='realid'
                     )

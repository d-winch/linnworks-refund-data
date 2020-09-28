import pandas as pd
import requests
import io
import os
from lxml import etree

def process_refunds(df):
    
    refunds = {
        'SR Sales': 0,
        'SR Ship': 0,
        'Z Sales': 0,
        'Z Ship': 0,
        'E Sales': 0,
        'E Ship': 0
    }

    # Get values from data's first row
    currency = df.iloc[0]['Currency']
    month = df.iloc[0]['CreateDate'].to_pydatetime().month
    year = df.iloc[0]['CreateDate'].to_pydatetime().year
    mmyy = f"{str(month).rjust(2, '0')}{str(year)[-2:]}"

    # Get exchange rate
    rate = get_rate(mmyy, currency)
    print(f'{currency} Rate: ', rate)

    # For each row in this dataframe (subsource and currency)
    for _, row in df.iterrows():
        
        # Get data
        category = row['ProductCategory']
        tax_rate = row['Tax Rate']
        reason = row['Reason']
        amount = float(row['Amount'])

        # Add amount to correct dict key
        if tax_rate == 20 and category == 'fullVat' and not reason == 'Shipping costs refund':
            refunds['SR Sales'] += amount
        if tax_rate == 20 and category == 'fullVat' and reason == 'Shipping costs refund':
            refunds['SR Ship'] += amount
        if tax_rate == 20 and category == 'ZeroVat' and not reason == 'Shipping costs refund':
            refunds['Z Sales'] += amount
        if tax_rate == 20 and category == 'ZeroVat' and reason == 'Shipping costs refund':
            refunds['Z Ship'] += amount
        if tax_rate == 0 and not reason == 'Shipping costs refund':
            refunds['E Sales'] += amount
        if tax_rate == 0 and reason == 'Shipping costs refund':
            refunds['E Ship'] += amount

    # Multiply each dict value by the exchange rate
    refunds.update((x, y*rate) for x, y in refunds.items())
    
    return rate, refunds

def get_category(row):
    # Get parent sku and category data (fullVat, ZeroVat, or null)
    sku = row['SKU'][:5]
    category = row['ProductCategory']

    # If not null, we know the category
    if not pd.isnull(category):
        return category

    # If parent sku in these known values, it is fullVat
    if sku in [
            "A112F", "AV105", "BE089",
            "BE104", "GD001", "GD014",
            "GD056", "GD057", "GD058",
            "GD072", "GD076", "JH030",
            "PR154", "SK271", "SK303",
            "SS024", "SS026", "SS028",
            "SS030", "SS100", "YK001",
            "YK180", "HVW10"]:
        return 'fullVat'

    # If parent sku in these known values, it is ZeroVat
    if sku in [
            "00000", "A105B", "AC01J",
            "B111B", "BZ002", "BZ011",
            "BZ031", "BZ032", "GD01B",
            "GD56B", "GD57B", "GD58B",
            "JH01J", "JH09J", "JH43J",
            "JH51J", "JH53J", "LW02T",
            "LW25T", "SM271", "SS007",
            "SS031"]:
        return 'ZeroVat'

    ###########################
    #      Special cases      #
    #  Al's Halloween designs #
    if row['SKU'][:2] == 'CG':
        return 'fullVat'
    #   Screenprint designs   #
    if row['SKU'][:3] in ('JR0', 'JR1'):
        return 'fullVat'
    ###########################

    # Else, we need to ask the user for the category and return it
    category_input = ''
    while category_input not in ['fullVat', 'ZeroVat']:
        category_input = input(f'Please enter a category for {sku}\nEither fullVat or ZeroVat: ')
    return category_input

def get_rate(mmyy, currency):
    """ Returns the exchange rate for the given month and year for this currency """
    if currency == 'GBP':
        return 1
    response = requests.get(f'http://www.hmrc.gov.uk/softwaredevelopers/rates/exrates-monthly-{mmyy}.xml')
    parser = etree.XMLParser(recover=False)
    #root = etree.fromstring(response.content.decode('utf-8').replace(" Period=""01/Jun/2019 to 30/Jun/2019""",''), parser=parser)
    rate = root.xpath(f"//exchangeRateMonthList/exchangeRate[currencyCode='{currency}']/rateNew")[0]
    rate = float(rate.text)
    return 1/rate


if __name__ == '__main__':

    # Read data and drop duplicate rows
    df = pd.read_csv('All.csv')
    df.drop_duplicates(subset="pkRefundRowId",
                       keep=False, inplace=True)
    
    # Parse date - Try with/without second data
    try:
        df['CreateDate'] = pd.to_datetime(df['CreateDate'], format='%d/%m/%Y %H:%M:%S', utc=True)
    except:
        df['CreateDate'] = pd.to_datetime(df['CreateDate'], format='%d/%m/%Y %H:%M', utc=True)

    # Set the category for any row with null data
    df['ProductCategory'] = df.apply(get_category, axis=1)

    # Create empty array to store refund data to write later
    refund_array = []
    
    # For each subsource and currency in the data
    for subsource in set(df['SubSource']):
        for currency in set(df.loc[df['SubSource'] == subsource, 'Currency']):
            print(subsource, currency)
            # Get the rows with this subsource and currency
            refunds = df.loc[(df['SubSource'] == subsource) & (df['Currency'] == currency)]
            print(refunds)
            # Process refund data, add to refund array, and write data to a csv
            rate, refund_data = process_refunds(refunds)
            refund_array.append((subsource, currency, rate, refund_data))
            refunds.to_csv(f"{subsource} {currency}.csv", index=False)

    # Get the month directory
    directory = os.path.basename(os.getcwd())

    # Sort the array for consistency
    refund_array.sort()
    
    # Write refund data to file
    with open(f"refunds {directory}.txt", "w") as f:
        for subsource, currency, rate, refunds in refund_array:
            f.write(f'{subsource} {currency} -> GBP\n')
            f.write(f'Rate: {rate}\n')
            for k, v in refunds.items():
                f.write(f'{k}: {v:.2f}\n')
            f.write('\n')


import os

import pandas as pd
import psycopg2 as ps

from secret import *

from os.path import expanduser
import requests
import plotly
from subprocess import call

pd.set_option('display.height', 1000)
pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)


def get_first_bday_of_month(mnth=None, yr=None):
    '''
    Return the first business day of the current month if no variables provided
    Return the first business day of the month and year provided if variables provided

    Tests:
    In [188]: config.get_first_bday_of_month(12,2015)
    Out[188]: datetime.date(2015, 12, 1)

    In [189]: config.get_first_bday_of_month(11,2015)
    Out[189]: datetime.date(2015, 11, 2)

    In [190]: config.get_first_bday_of_month(10,2015)
    Out[190]: datetime.date(2015, 10, 1)

    In [191]: config.get_first_bday_of_month(1,2016)
    Out[191]: datetime.date(2016, 1, 4)

    In [192]: config.get_first_bday_of_month(8,2015)
    Out[192]: datetime.date(2015, 8, 3)
    :param mnth:
    :param yr:
    :return:
    '''
    from calendar import monthrange
    from pandas.tseries.holiday import USFederalHolidayCalendar
    from pandas.tseries.offsets import CustomBusinessDay
    if yr is None or mnth is None:
        yr = pd.datetime.now().year if pd.datetime.now().month != 1 else pd.datetime.now().year - 1
        mnth = pd.datetime.now().month - 1 if pd.datetime.now().month != 1 else 12
    else:
        yr = yr if mnth != 1 else yr - 1
        mnth = mnth - 1 if mnth != 1 else 12
    end_last = monthrange(yr, mnth)
    end_last = pd.Timestamp('%s/%s/%s' % (mnth, end_last[1], yr)).date()
    cal = USFederalHolidayCalendar()
    holidays = cal.holidays(start=end_last - pd.tseries.offsets.Day(60),
                            end=end_last + pd.tseries.offsets.Day(60)).to_pydatetime()
    bday_cus = CustomBusinessDay(holidays=holidays)
    return (end_last + bday_cus).date()


def get_last_weekday_value(weekday):  # Mon = 0, Sun = 6
    '''
    If you want to get the last Friday, use weekday=4, last Thurs = 3, Sun = 6
    '''
    day_minus = pd.datetime.today().weekday() + weekday - 1
    return (pd.datetime.today() - pd.tseries.offsets.Day(day_minus)).date()

def plotly_retries(url):
    """
    Can mostly get away without using this as we
    implemented this type of behavior directly into Plotly,
    but it's not a bad check to handle Plotly's servers
    being wonky
    :param url:
    :return:
    """
    import logging
    logging.basicConfig(filename='%splotly.log' % excel_folder, level=logging.DEBUG,
                        format='%(asctime)s %(message)s')
    attempts = 0
    logging.debug('status code == {}, embed status code == {}, {}, attempt {}'.format(
            requests.get(url.split('?')[0] + '.png?' + url.split('?')[1]).status_code,
            requests.get(url.split('?')[0] + '.embed?' + url.split('?')[1]).status_code,
            url, attempts))
    while (requests.get(url.split('?')[0] + '.png?' + url.split('?')[1]).status_code == 404 or
                   requests.get(url.split('?')[0] + '.embed?' + url.split('?')[1]).status_code == 404):
        logging.debug('404 on {}, attempt {}'.format(url, attempts))
        attempts += 1
        if attempts == 7:
            break
        # attempt to add secret sharing permissions again
        url = plotly.plotly.plotly.add_share_key_to_url(url.split('?')[0])
    url_split = url.split('?')
    return url_split




'''
##############
#
#Getting files and putting files to dropbox.
#
##############
'''


def put_dropbox(read_loc, file_loc, overwrite=False, remove=True):
    """
    This puts a specified file into a specified dropbox location
    Does a try/except check to catch any problems with Dbox servers
    or your connection
    :param read_loc: Location of file to put on dropbox
    :param file_loc: Dropbox folder location
    :param overwrite: Overwrite any existing Dropbox files of this location
    :param remove: Remove file from your machinr after
    :return:
    """
    import auth_dropbox
    ct = 1
    while ct < 3:
        try:
            client = auth_dropbox.auth()
            f = open(read_loc, 'rb')
            response = client.put_file(file_loc, f, overwrite=overwrite)
            if remove:
                os.remove(read_loc)
            return response
        except Exception, e:
            import time
            time.sleep(20)
            ct += 1


def get_previous_dropbox(folder):
    '''
    #######
    #
    #Format needs to be Name_Of_File MM_DD_YYYY.xlsx
    #This actually finds the latest file that we uploaded and uses it as the "latest file"
    #
    #######
    '''
    import auth_dropbox
    client = auth_dropbox.auth()
    mst = {}
    contents = client.metadata(folder)['contents']
    for content in contents:
        retries = 0
        while retries < 3:
            try:
                mst[pd.Timestamp(content['path'].split('/')[-1].split(' ')[-1].split('.xlsx')[0].replace('_', '/'))] = {
                    'Path': content['path']}
                retries = 3
            except Exception, e:
                print e
                retries += 1
    # Grabs a dataframe of file names with the index as the date on the file(not modified, the actual title)
    df = pd.DataFrame.from_dict(mst, orient='index')
    df.sort_index(inplace=True)
    file_ = client.get_file(df['Path'].ix[-1])
    lst = pd.read_excel(file_, index_col=0)
    return lst


def get_dropbox(path):
    import auth_dropbox
    client = auth_dropbox.auth()
    # TODO: Check for csv or .xlsx
    return pd.read_csv(client.get_file(path))


def format_percent(item):
    return '{:.1%}'.format(item)


def format_money(item, locale='us'):
    if locale == 'uk':
        return u'\xA3{:,.0f}'.format(item)
    else:
        return '${:,.0f}'.format(item)


def format_money_gbp(item, locale='uk'):
    if locale == 'uk':
        return u'\xA3{:,.0f}'.format(item)
    else:
        return '${:,.0f}'.format(item)


def format_money_w_decimal(item, locale='us'):
    if locale == 'uk':
        return u'\xA3{:,.0f}'.format(item)
    else:
        return '${:,.2f}'.format(item)


def format_number(item):
    return '{:,.0f}'.format(item)


def format_number_w_decimal(item):
    return '{:,.1f}'.format(item)


__author__ = "Chase Schwalbach"
__credits__ = ["Chase Schwalbach"]
__version__ = "1.0"
__maintainer__ = "Chase Schwalbach"
__email__ = "chase.schwalbach@avant.com"
__status__ = "Production"

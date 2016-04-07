import gspread
import pandas as pd
from oauth2client.client import SignedJwtAssertionCredentials

import config
import send_email


class AliveEmailer(object):
    def __init__(self):
        self.wks_names = {'Adopter Information': 0,
                          '1 week': 1,
                          '1 month': 2,
                          '3 month': 3,
                          '1 year': 4}
        scope = ['https://spreadsheets.google.com/feeds']
        credentials = SignedJwtAssertionCredentials(config.drive_details['client_email'],
                                                    config.drive_details['private_key'], scope)
        gc = gspread.authorize(credentials)
        # Open a worksheet from spreadsheet with one shot
        # tit = "ALIVE Rescue Little Barn Volunteer Schedule"
        tit = 'Adoption Follow Up Database'
        self.wkbk = gc.open(tit)
        self._check_for_emails()
        self.get_sheet_by_name(sheetname=self.wks_names['1 month'])

    def main(self):
        for idx, row in self.adopter.iterrows():
            adopt_date = row['Adoption Date']
            if row['1 week'] == 'TRUE':  # Need to send, if after 1 week
                if pd.datetime.today() >= adopt_date + pd.tseries.offsets.Week(1):
                    print 'week'
                    print row
            if row['1 month'] == 'TRUE':
                if pd.datetime.today() >= adopt_date + pd.tseries.offsets.Day(30):
                    print 'month'
                    print row
            if row['3 month'] == 'TRUE':
                if pd.datetime.today() >= adopt_date + pd.tseries.offsets.Day(90):
                    print '3 month'
                    print row
            if row['1 year'] == 'TRUE':
                if pd.datetime.today() >= adopt_date + pd.tseries.offsets.Day(365):
                    print 'year'
                    print row
    def _check_for_emails(self):
        sheetname = 'Adopter Information'
        wks = self.wkbk.get_worksheet(self.wks_names[sheetname])
        self.adopter = pd.DataFrame(wks.get_all_records())
        self.adopter['Adoption Date'] = pd.to_datetime(self.adopter['Adoption Date'])

    def _send_new_email(self, to, subject):
        self.eh = send_email.EmailHandler(to=to, subject=subject, from_='ALIVE Rescue')
        self.eh.add_random_text('''<br><br>Sarah Brewster<br>
                              Director of Adoptions and Foster''')
        self.eh.send_email('sarahb.aliverescue@gmail.com', 'sjb#9405')

    def get_sheet_by_name(self, sheetname):
        wks = self.wkbk.get_worksheet(sheetname)
        print wks.get_all_records()
        return wks


if __name__ == '__main__':
    AliveEmailer()

__author__ = "Chase Schwalbach"
__credits__ = ["Chase Schwalbach"]
__version__ = "1.0"
__maintainer__ = "Chase Schwalbach"
__email__ = "chase.schwalbach@avant.com"
__status__ = "Production"

import gspread
import pandas as pd
from oauth2client.client import SignedJwtAssertionCredentials
# Config relies on secret.py
import config
# send_email uses secret.py for logins
import send_email


class AliveEmailer(object):
    def __init__(self):
        """
        Script to automate e-mails for ALIVE
        """
        self.wks_names = {'1 week': 1,
                          '1 month': 2,
                          '3 month': 3,
                          '1 year': 4}
        scope = ['https://spreadsheets.google.com/feeds']
        credentials = SignedJwtAssertionCredentials(config.drive_details['client_email'],
                                                    config.drive_details['private_key'], scope)
        gc = gspread.authorize(credentials)
        # Open a worksheet from spreadsheet with one shot
        self.summary_email = send_email.EmailHandler(to=config.master_email, subject='E-mails being sent out today')
        tit = 'Adoption Follow Up Database'
        self.wkbk = gc.open(tit)
        self.adopter = self._get_adopter_info()
        self.wks_text = self._get_email_text()

    def main(self):
        """
        Running through the Google sheet, checking
        for rows (adopters) that have true, which means
        they need to send an email still.
        """
        for idx, row in self.adopter.iterrows():
            adopt_date = row['Adoption Date']
            if row['1 week'] == 'TRUE':  # Need to send, if after 1 week
                if pd.datetime.today() >= adopt_date + pd.tseries.offsets.Week(1):
                    # This Adopter needs to send out the weekly update
                    self._send_update_email(idx, row, '1 week')
            if row['1 month'] == 'TRUE':
                if pd.datetime.today() >= adopt_date + pd.tseries.offsets.Day(30):
                    self._send_update_email(idx, row, '1 month')
            if row['3 month'] == 'TRUE':
                if pd.datetime.today() >= adopt_date + pd.tseries.offsets.Day(90):
                    self._send_update_email(idx, row, '3 month')
            if row['1 year'] == 'TRUE':
                if pd.datetime.today() >= adopt_date + pd.tseries.offsets.Day(365):
                    self._send_update_email(idx, row, '1 year')

    def _send_update_email(self, idx, row, type_email):
        """
        When something needs to be sent, this will add to a summary email and send
        :param idx:
        :param row:
        :param type_email: key to self.wks_names / self.wks_text
        :return:
        """
        pass

    def _get_adopter_info(self):
        wks = self.wkbk.get_worksheet(0)
        adopter = pd.DataFrame(wks.get_all_records())
        adopter['Adoption Date'] = pd.to_datetime(adopter['Adoption Date'])
        return adopter

    def _send_new_email(self, to, subject):
        self.eh = send_email.EmailHandler(to=to, subject=subject, from_='ALIVE Rescue', user=config.gmail_user,
                                          pwd=config.gmail_pwd)
        self.eh.add_random_text('''<br><br>Sarah Brewster<br>
                              Director of Adoptions and Foster''')
        self.eh.send_email()

    def _get_email_text(self):
        """
        Get the text for each email that we'll be using
        :return:
        """
        master = {}
        for key, val in self.wks_names.iteritems():
            wks = self._get_sheet_by_name(sheetname=val)
            txt = wks.get_all_records()
            master[key] = txt[0]
        return master

    def _get_sheet_by_name(self, sheetname):
        wks = self.wkbk.get_worksheet(sheetname)
        return wks


if __name__ == '__main__':
    ae = AliveEmailer()
    ae.main()

__author__ = "Chase Schwalbach"
__credits__ = ["Chase Schwalbach"]
__version__ = "1.0"
__maintainer__ = "Chase Schwalbach"
__email__ = "chase.schwalbach@avant.com"
__status__ = "Production"

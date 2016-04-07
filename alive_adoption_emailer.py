import gspread
import pandas as pd
from oauth2client.client import SignedJwtAssertionCredentials
# Config relies on secret.py
import config
# send_email uses secret.py for logins
import send_email
from send_email import EmailHandler


class AliveEmailer(object):
    def __init__(self):
        """
        Script to automate e-mails for ALIVE
        """
        scope = ['https://spreadsheets.google.com/feeds']
        credentials = SignedJwtAssertionCredentials(config.drive_details['client_email'],
                                                    config.drive_details['private_key'], scope)
        gc = gspread.authorize(credentials)
        # Open a worksheet from spreadsheet with one shot
        self.summary_email = EmailHandler(to=config.master_email, subject='E-mails being sent out today')
        self.summary_email.add_random_text('<ul>')
        tit = 'Adoption Follow Up Database'
        self.wkbk = gc.open(tit)
        # Now get the names, and find their index for parsing later
        self.wks_names = {wks.title: ix for ix, wks in enumerate(self.wkbk.worksheets())}
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
        self.summary_email.send_email()

    def _send_update_email(self, idx, row, type_email):
        """
        When something needs to be sent, this will add to a summary email and send
        :param idx:
        :param row:
        :param type_email: key to self.wks_names / self.wks_text
        :return:
        """
        # Send the e-mail to the person and Sarah
        email_txt = self.wks_text[type_email].format(Adopter_First_Name=row['Adopter First Name'],
                                                     Pet_NAME=row['PET Name'])
        email_subject = 'ALIVE Rescue follow up!'
        send_email.send_email(to_user='schwallie@gmail.com', SUBJECT=email_subject, TEXT=email_txt, FROM='ALIVE Rescue')
        # TODO: Need to update the field in the Google Sheet Now!
        # Now send a summary email to Sarah
        self.summary_email.add_random_text('<li>%s: %s</li>' % (row['PET Name'], type_email))

    def _get_adopter_info(self):
        wks = self.wkbk.get_worksheet(self.wks_names['Adopter Information'])
        adopter = pd.DataFrame(wks.get_all_records())
        adopter['Adoption Date'] = pd.to_datetime(adopter['Adoption Date'])
        return adopter

    def _add_footer(self, eh):
        eh.add_random_text('''<br><br>Sarah Brewster<br>
                              Director of Adoptions and Foster''')

    def _get_email_text(self):
        """
        Get the text for each email that we'll be using
        :return:
        """
        master = {}
        for key, val in self.wks_names.iteritems():
            print key
            if '1' in key or '3' in key:  # 1 week, 1 month, 3 month, 1 year
                wks = self.wkbk.get_worksheet(val)
                txt = wks.get_all_values()[0][0]
                master[key] = txt
        return master


if __name__ == '__main__':
    ae = AliveEmailer()
    ae.main()

__author__ = "Chase Schwalbach"
__credits__ = ["Chase Schwalbach"]
__version__ = "1.0"
__maintainer__ = "Chase Schwalbach"
__email__ = "chase.schwalbach@avant.com"
__status__ = "Beta"

# system and cli interface
import csv
import getopt
import json
import sys

# dependencies
from oauth2client.client import SignedJwtAssertionCredentials
import gspread


def num_to_letter(num):
    letter = chr(num % 26 - 1 + ord('A'))
    count = num / 26
    prepend = chr(count % 26 - 1 + ord('A')) if count else ""
    return "%s%s" % (prepend, letter)


def get_range(col_end, row_end):
    return "A1:%s%d" % (col_end, row_end)


class csv2GSheet():
    """Class wrapper to upload a csv to a google spreadsheet"""

    scope = ['https://spreadsheets.google.com/feeds']

    def __init__(self, oauth=None):
        self.cli = self._authorize_oauth(oauth)
        self.cli.login()

    def client(self):
        return self.cli

    def _authorize_oauth(self, creds):
        """perform oauth dance and return validated client"""
        credentials = SignedJwtAssertionCredentials(creds["client_email"],
                                                    creds["private_key"],
                                                    self.scope)
        gc = gspread.authorize(credentials)
        return gc

    def uploadCSV(self, fcsv, destination, wksheet=None):
        """Writes open file csv to destination spreadsheet / worksheet"""
        csv_reader = csv.reader(fcsv)
        sh = self.client().open(destination)
        current_wksheets = sh.worksheets()
        wk = None
        for curr_sheet in current_wksheets:
            if wksheet == curr_sheet.title:
                wk = curr_sheet
                break
        if not wk:
            wk = sh.add_worksheet(title=(wksheet if wksheet else "csv2GSheet"), rows="100", cols="20")
        index = 0
        rows = []
        for row in csv_reader:
            rows.append(row)
            index += 1
        if not rows:
            return
        col_length = len(rows[0])
        cell_range_label = get_range(num_to_letter(col_length), index)
        cell_list = wk.range(cell_range_label)
        i = 0
        j = 0
        for c in cell_list:
            c.value = rows[i][j]
            j += 1
            if j % (col_length) == 0:
                j = 0
                i += 1
        wk.update_cells(cell_list)


def handleUsageError():
    """Print usage error"""
    print 'python csv2gsheet.py --creds [creds] --target [file] --dest_file [file] --dest_wksheet [wksheet]'
    sys.exit(2)


def main():
    """Formulate cli args and send target CSV to target GSheet

    :: creds - a json file of service account oauth creds
    :: target - the CSV file to upload
    :: dest_file - the spreadsheet to put the csv file
    :: dest_wksheet - the worksheet in the spreadsheet to put the csv file (not required)

    """
    try:
        opts, args = getopt.getopt(sys.argv[1:], "", ["creds=", "target=", "dest_file=", "dest_wksheet="])
    except (getopt.error):
        return handleUsageError()

    creds = ""
    target = ""
    dest_file = ""
    dest_wksheet = ""
    for k, v in opts:
        if k == "--creds":
            creds = v
        elif k == "--target":
            target = v
        elif k == "--dest_file":
            dest_file = v
        elif k == "--dest_wksheet":
            dest_wksheet = v

    if (creds == "") or (target == "") or (dest_file == ""):
        return handleUsageError()

    credentials = json.load(open(creds))
    cli = csv2GSheet(oauth=credentials)

    with open(target, "r") as infile:
        print "writing file to target"
        cli.uploadCSV(infile, dest_file, wksheet=dest_wksheet)


if __name__ == "__main__":
    main()

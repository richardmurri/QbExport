#!/usr/bin/python

from __future__ import with_statement
from decimal import Decimal, InvalidOperation
import sys, os, collections, Tkinter, tkFileDialog, csv


class LogWriter(object):
    def __init__(self, *logs):
        self.logs = logs

    def write(self, message):
        for log in self.logs:
            log(message)

class MessageWindow(object):
    """ Displays error messages using tkinter """
    def __init__(self):

        # set up messages window
        Tkinter.Label(text='Messages:').pack()

        # add right scrollbar
        s = Tkinter.Scrollbar()
        s.pack(side=Tkinter.RIGHT, fill=Tkinter.Y)

        # add bottom scrollbar
        v = Tkinter.Scrollbar(orient=Tkinter.HORIZONTAL)
        v.pack(side=Tkinter.BOTTOM, fill=Tkinter.X)

        # add list box
        lbox = Tkinter.Listbox(width=80, height=12)
        lbox.pack(expand=True, fill=Tkinter.BOTH, side=Tkinter.TOP)

        # pack the components
        s.config(command=lbox.yview)
        lbox.config(yscrollcommand=s.set)
        v.config(command=lbox.xview)
        lbox.config(xscrollcommand=v.set)

        self.lbox = lbox

    def insert(self, text):
        self.lbox.insert(Tkinter.END, text)

    def show(self):
        Tkinter.mainloop()

class UI(object):
    def __init__(self):

        # start tk and ask it to not show a window
        self.root = Tkinter.Tk()
        self.root.withdraw()

    def get_filename(self):
        filename = tkFileDialog.askopenfilename(filetypes=[('Quickbooks Export', '*.csv')])
        self.root.deiconify()
        return filename

# Make the ui a singleton object
UI = UI()

class ParseError(Exception): pass
class NotImplementedError(Exception): pass
class VoidError(Exception): pass
class ColumnsInvalidError(Exception): pass


class FileParser(object):
    def __init__(self, filename, log_writer):
        self.filename = filename
        self.log_writer = log_writer
        self.currentAccount = [] # a stack, keeps track of the current account while parsing

    def get_decimal(self, value):
        """ Return a decimal value from a string """
        try:
            amount = Decimal(value)
        except InvalidOperation:
            amount = Decimal(0)
        return amount

    def check_columns(self, columnArray):
        """ Check that all needed columns of data are available """
        for column in ('AccountName', 'Trans #', 'Type', 'Split', 'Date', 'Name', 'Memo'):
            if column not in columnArray:
                raise ColumnsInvalidError('The necessary column "%s" was not found in the input file' % column)

        if 'Amount' not in columnArray and ('Debit' not in columnArray or 'Credit' not in columnArray):
            raise ColumnsInvalidError('The necessary column "Amount" was not found in the input file')

    def get_data_row(self, line, columnArray):
        """ Parse a row of the quickbooks data and return the result as a map """

        # if this is a row specifying the account then push or pop the account
        if line[0]:
            if line[0].startswith('Total'):
                self.currentAccount.pop()
            elif line[0] == 'TOTAL':
                return None
            else:
                self.currentAccount.append(line[0])
            return None
        else:
            # make a map
            new = {}
            for i, data in enumerate(line):
                new[columnArray[i]] = data

            # Fix data being read
            if 'Credit' in columnArray and 'Debit' in columnArray:
                credit = -self.get_decimal(new['Credit'])
                debit = self.get_decimal(new['Debit'])
                new['Amount'] = credit if credit else debit
            else:
                new['Amount'] = self.get_decimal(new['Amount'])

            new['AccountName'] = list(self.currentAccount)
        return new

    def parse_file(self):
        """ Read a csv file and parse the data into a list """
        transactions = []

        with open(self.filename) as f:

            reader = csv.reader(f)

            # parse header columns
            columnArray = reader.next()
            columnArray[0] = 'AccountName'

            # check to make sure we have all the needed columns
            try:
                self.check_columns(columnArray)
            except ColumnsInvalidError, e:
                self.log_writer.write(e)
                raise ParseError()

            transaction = []
            try:
                # loop through data
                for line in reader:
                    transaction = self.get_data_row(line, columnArray)
                    if transaction is not None:
                        transactions.append(transaction)
            except csv.Error, e:
                self.log_writer.write('Could not parse file (line number %s)' % reader.line_num)
                raise ParseError()

        return transactions


class IIFGenerator(object):
    file_start_tpl = '!TRNS\tTRNSID\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tDOCNUM\tMEMO\tCLEAR\tTOPRINT\tADDR1\tADDR2\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\n' \
                     '!SPL\tSPLID\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tDOCNUM\tMEMO\tCLEAR\tQNTY\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\n' \
                     '!ENDTRNS\n'
    trans_tpl = 'TRNS\t\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\n'
    split_tpl = 'SPL\t\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\n'
    trans_end_tpl = 'ENDTRNS\n'

    def __init__(self, iif_filename, log_writer):
        self.iif_filename = iif_filename
        self.log_writer = log_writer

    def numeric_compare(self, x, y):
        return cmp(int(x), int(y))

    def generate(self, transactions):
        # don't handle any type of transaction other than check and deposit
        typesNotImplemented = [set(), []]
        voidErrors = []

        # any transaction that fails will be shown to the user
        failures = []

        # index transactions by id
        transMap = collections.defaultdict(list)
        for trans in transactions:
            transId = trans['Trans #']
            transMap[transId].append(trans)

        with open(self.iif_filename, 'w') as f:
            f.write(self.file_start_tpl)

            for transId, transactions in transMap.iteritems():
                try:
                    trans, splits = self.decipher_transactions(transactions)

                    self.write_transaction(trans, splits, f)
                except ParseError, e:
                    self.log_writer.write('%s - %s' % (transId, e))
                except NotImplementedError, e:
                    typesNotImplemented[0].add(str(e))
                    typesNotImplemented[1].append(transId)
                except VoidError, e:
                    voidErrors.append(transId)

        if typesNotImplemented[0]:
            typesNotImplemented[1].sort(cmp=self.numeric_compare)
            self.log_writer.write('This program in not yet able to translate any of these types %s.  These transactions were not added %s' % (list(typesNotImplemented[0]), typesNotImplemented[1]))
        if voidErrors:
            voidErrors.sort(cmp=self.numeric_compare)
            self.log_writer.write('Some transactions could not be written.  This could be caused by a transaction amount of $0.00 (i.e. a voided check).  These transactions were not added %s' % voidErrors)


    def decipher_transactions(self, transactions):
        outgoing = ['Check']
        incoming = ['Deposit']

        tType = transactions[0]['Type']

        if tType in outgoing:
            tranList = [ val for val in transactions if val['Amount'] < 0 ]
            splits =  [ val for val in transactions if val['Amount'] > 0 ]
        elif tType in incoming:
            tranList = [ val for val in transactions if val['Amount'] > 0 ]
            splits =  [ val for val in transactions if val['Amount'] < 0 ]
        elif transactions[0]['Amount'] == 0:
            tranList = [ val for val in transactions if val['Split'] == 'void' ]
            tranList = [ val for val in transactions if val['Split'] != 'void' ]
        else:
            raise NotImplementedError(tType)

        if len(tranList) != 1:
            raise VoidError('There was a problem deciphering the data, possibly because the transaction amount is $0.00')

        trans = tranList[0]

        # if there are no splits try to recreate them
        if len(splits) == 0:
            if trans['Split'] == '-SPLIT-':
                raise ParseError('Can not resolve the split values')
            elif trans['Split'] == '':
                raise ParseError('There are no splits specified')
            else:
                spl = dict(trans)
                spl['Amount'] = -1 * spl['Amount']
                spl['AccountName'], spl['Split'] = [ spl['Split'] ], spl['AccountName'][-1]
                splits.append(spl)

        # check to make sure splits add up
        tranSum = -Decimal(trans['Amount'])
        splitSum = sum([ val['Amount'] for val in splits ])
        if not tranSum == splitSum:
            raise ParseError('The sum of the splits does not equal the total of the transaction')

        return trans, splits

    def write_transaction(self, trans, splits, f):
        # create the main transaction
        accountName = ':'.join(trans['AccountName'])
        tType = trans['Type']
        date = trans['Date']
        name = trans['Name']
        amount = trans['Amount']
        memo = trans['Memo']
        num = trans['Num'] if trans['Num'] else ''
        cleared = 'N'
        toPrint = 'N'
        address1 = ''
        address2 = ''

        trans = self.trans_tpl % (tType, date, accountName, name, amount, num, memo, cleared, toPrint, address1, address2)
        f.write(trans)

        # create the splits
        for spl in splits:
            splAccount = ':'.join(spl['AccountName'])
            splAmount = spl['Amount']
            splCleared = 'N'
            splString = self.split_tpl % (tType, date, splAccount, name, splAmount, num, memo, splCleared)
            f.write(splString)

        f.write(self.trans_end_tpl)


def get_iif_filename(filename):
    # keep the same filename, just swap the extension
    dir_, file_ = os.path.split(filename)
    file_, ext = file_.split('.')
    return os.path.join(dir_, file_ + '.iif')

def get_log_writer(filename):
    dir_, file_ = os.path.split(filename)
    file_, ext = file_.split('.')
    new_filename = os.path.join(dir_, file_ + '.log')
    open_file = open(new_filename, 'w')
    def logfile_writer(message):
        open_file.write('%s\n\n' % message)
    return logfile_writer


if __name__ == '__main__':

    # get file to open
    if len(sys.argv) > 1:
        filename = sys.argv[1]
    else:
        filename = UI.get_filename()

    # get error message window
    message_window = MessageWindow()

    # if not file was specified exit
    if not filename:
        message_window.insert('No file to convert, quitting ...')
        message_window.show()
        sys.exit(1)

    # create logfile writer function
    logfile_writer = get_log_writer(filename)

    # create log_writer instance that will write log messages to both the message window and the log file
    log_writer = LogWriter(message_window.insert, logfile_writer)

    # find filename of the iif file to write
    iif_filename = get_iif_filename(filename)

    try:
        # find the files transactions and index by i
        transactions = FileParser(filename, log_writer).parse_file()

        # generate the iif file
        IIFGenerator(iif_filename, log_writer).generate(transactions)

    except ParseError, e:
        message_window.show()
        sys.exit(1)

    # show log messages
    log_writer.write('File created: %s' % iif_filename)
    message_window.show()

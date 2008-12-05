#!/usr/bin/python

from __future__ import with_statement
from decimal import Decimal, InvalidOperation
import sys, os, collections, Tkinter, tkFileDialog, csv

def getFileName():
    """ Get the name of the file to open from the arguments or from the user """
    if len(sys.argv) > 1:
        return sys.argv[1]
    else:
        # ask for filename
        root = Tkinter.Tk()
        root.withdraw()
        file = tkFileDialog.askopenfilename(filetypes=[('Quickbooks Export', '*.csv')])
        root.deiconify()
        return file

def numericCompare(x, y):
    return cmp(int(x), int(y))

def getFileToWrite(file):
    # keep the same filename, just swap the extension
    dir, file = os.path.split(file)
    file, ext = file.split('.')
    return os.path.join(dir, file + '.iif')

def getLogFile(file):
    dir, file = os.path.split(file)
    file, ext = file.split('.')
    return os.path.join(dir, file + '.log')

class MessageWindow():
    """ Displays error messages using tkinter """
    def __init__(self, file):
        # set up messages window
        Tkinter.Label(text='Messages:').pack()
        s = Tkinter.Scrollbar()
        s.pack(side=Tkinter.RIGHT, fill=Tkinter.Y)
        v = Tkinter.Scrollbar(orient=Tkinter.HORIZONTAL)
        v.pack(side=Tkinter.BOTTOM, fill=Tkinter.X)
        lbox = Tkinter.Listbox(width=80, height=12)
        lbox.pack(expand=True, fill=Tkinter.BOTH, side=Tkinter.TOP)
        s.config(command=lbox.yview)
        lbox.config(yscrollcommand=s.set)
        v.config(command=lbox.xview)
        lbox.config(xscrollcommand=v.set)
        self.lbox = lbox
        if file:
            logFile = getLogFile(file)
            self.log = open(logFile, 'w')
        else:
            self.log = sys.stdout

    def insert(self, text):
        self.lbox.insert(Tkinter.END, text)
        self.log.write(text + '\n')

    def show(self):
        Tkinter.mainloop()
        if self.log != sys.stdout:
            self.log.close()
        sys.exit(0)

class ParseError(Exception):
    pass

class NotImplementedError(Exception):
    pass

class VoidError(Exception):
    pass

class ColumnsInvalidError(Exception):
    pass

def checkColumns(columnArray):
    """ Check that all needed columns of data are available """
    for column in ('AccountName', 'Trans #', 'Type', 'Split', 'Date', 'Name', 'Memo'):
        if column not in columnArray:
            raise ColumnsInvalidError('The necessary column "%s" was not found in the input file' % column)

    if 'Amount' not in columnArray and ('Debit' not in columnArray or 'Credit' not in columnArray):
        raise ColumnsInvalidError('The necessary column "Amount" was not found in the input file')

def getDecimal(value):
    """ Return a decimal value from a string """
    try:
        amount = Decimal(value)
    except InvalidOperation:
        amount = Decimal(0)
    return amount

# use a stack to keep track of the current account
currentAccount = []
def getDataRow(line, columnArray):
    """ Parse a row of the quickbooks data and return the result as a map """
    global currentAccount

    # if this is a row specifying the account then push or pop the account
    if line[0]:
        if line[0].startswith('Total'):
            currentAccount.pop()
        elif line[0] == 'TOTAL':
            return None
        else:
            currentAccount.append(line[0])
        return None
    else:
        # make a map
        new = {}
        for i, data in enumerate(line):
            new[columnArray[i]] = data

        # Fix data being read
        if 'Credit' in columnArray and 'Debit' in columnArray:
            credit = -getDecimal(new['Credit'])
            debit = getDecimal(new['Debit'])
            new['Amount'] = credit if credit else debit
        else:
            new['Amount'] = getDecimal(new['Amount'])

        new['AccountName'] = list(currentAccount)
    return new

def parseFile(file, messageWindow):
    """ Read a csv file and parse the data into a list """
    transactions = []

    with open(file) as f:

        reader = csv.reader(f)

        # parse header columns
        columnArray = reader.next()
        columnArray[0] = 'AccountName'

        # check to make sure we have all the needed columns
        try:
            checkColumns(columnArray)
        except ColumnsInvalidError, e:
            messageWindow.insert(e)
            messageWindow.show()

        transaction = []
        try:
            # loop through data
            for line in reader:
                transaction = getDataRow(line, columnArray)
                if transaction is not None:
                    transactions.append(transaction)
        except csv.Error, e:
            messageWindow.insert('Could not parse file (line number %s)' % reader.line_num)
            messageWindow.show()

    return transactions

def indexById(transactions):
    """ Index the transactions by the Trans # """
    transMap = collections.defaultdict(list)

    for trans in transactions:
        transId = trans['Trans #']
        transMap[transId].append(trans)

    return transMap

def writeTransaction(trans, splits, f):
    global transTemplate
    global splitTemplate
    global transEnd

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

    trans = transTemplate % (tType, date, accountName, name, amount, num, memo, cleared, toPrint, address1, address2)
    f.write(trans)

    # create the splits
    for spl in splits:
        splAccount = ':'.join(spl['AccountName'])
        splAmount = spl['Amount']
        splCleared = 'N'
        splString = splitTemplate % (tType, date, splAccount, name, splAmount, num, memo, splCleared)
        f.write(splString)

    f.write(transEnd)


def generateIIF(newFile, transactions, messageWindow):
    global fileStart

    # don't handle any type of transaction other than check and deposit
    typesNotImplemented = [set(), []]
    voidErrors = []

    # any transaction that fails with be shown to the user
    failures = []

    with open(newFile, 'w') as f:
        f.write(fileStart)

        for transId, transactions in transactionMap.iteritems():
            try:
                trans, splits = decipherTransactions(transactions)

                writeTransaction(trans, splits, f)
            except ParseError, e:
                messageWindow.insert('%s - %s' % (transId, e))
            except NotImplementedError, e:
                typesNotImplemented[0].add(str(e))
                typesNotImplemented[1].append(transId)
            except VoidError, e:
                voidErrors.append(transId)

    if typesNotImplemented[0]:
        typesNotImplemented[1].sort(cmp=numericCompare)
        messageWindow.insert('This program in not yet able to translate any of these types %s.  These transactions were not added %s' % (list(typesNotImplemented[0]), typesNotImplemented[1]))
    if voidErrors:
        voidErrors.sort(cmp=numericCompare)
        messageWindow.insert('Some transactions could not be written.  This could be caused by a transaction amount of $0.00 (i.e. a voided check).  These transactions were not added %s' % voidErrors)

def decipherTransactions(transactions):
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
        splits = [ val for val in transactions if val['Split'] != 'void' ]
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

# handle data
fileStart = """\
!TRNS\tTRNSID\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tDOCNUM\tMEMO\tCLEAR\tTOPRINT\tADDR1\tADDR2\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t
!SPL\tSPLID\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tDOCNUM\tMEMO\tCLEAR\tQNTY\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t
!ENDTRNS
"""
transTemplate = """\
TRNS\t\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t
"""
splitTemplate = """\
SPL\t\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t
"""
transEnd = """\
ENDTRNS
"""

if __name__ == '__main__':

    # get file to open
    file = getFileName()

    # get error message window
    messageWindow = MessageWindow(file)

    if not file:
        messageWindow.insert('No file to convert, quitting ...')
        messageWindow.show()

    newFile = getFileToWrite(file)

    transactions = parseFile(file, messageWindow)
    transactionMap = indexById(transactions)

    generateIIF(newFile, transactionMap, messageWindow)

    messageWindow.insert('File created: %s' % newFile)
    messageWindow.show()

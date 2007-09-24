#!/usr/bin/python

from __future__ import with_statement
from decimal import Decimal, InvalidOperation
import sys, os, collections, Tkinter, tkFileDialog, csv, pprint

file = None

# find the file to import
try:
    file = sys.argv[1]
except IndexError:
    root = Tkinter.Tk()
    root.withdraw()
    file = tkFileDialog.askopenfilename(filetypes=[("Quickbooks Export", "*.csv")])
    root.deiconify()

# set up messages window
Tkinter.Label(text="Messages:").pack()
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

if not file:
    lbox.insert(Tkinter.END, "No file to convert, quitting ...")
    Tkinter.mainloop()
    exit()

transactionMap = collections.defaultdict(list)
with open(file) as f:

    reader = csv.reader(f)

    # parse header columns
    columnArray = reader.next()
    columnArray[0] = 'AccountName'
    
    currentAccount = []

    try:
        # loop through data
        for line in reader:
            # if the first column has data it is the account data
            if line[0]:
                if line[0].startswith('Total'):
                    currentAccount.pop()
                elif line[0] == 'TOTAL':
                    break;
                else:
                    currentAccount.append(line[0])
            else:
                # create a map of the data:  column->data
                newMap = {}
                for i, data in enumerate(line):
                    newMap[columnArray[i]] = data

                # fix data
                try:
                    amount = Decimal(newMap['Amount'])
                except InvalidOperation:
                    amount = Decimal(0)

                newMap['Amount'] = amount

                newMap['AccountName'] = list(currentAccount)
                transactionMap[newMap['Trans #']].append(newMap)

    except csv.Error, e:
        lbox.insert(Tkinter.END, 'Could not parse file (line number %s)' % reader.line_num)
        Tkinter.mainloop()
        exit()

# handle data
fileStart = """\
!TRNS\tTRNSID\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tDOCNUM\tMEMO\tCLEAR\tTOPRINT\tADDR1\tADDR2\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t
!SPL\tSPLID\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tDOCNUM\tMEMO\tCLEAR\tQNTY\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t
!ENDTRNS
"""
transEnd = """\
ENDTRNS
"""
transTemplate = """\
TRNS\t\t%s\t%s\t%s\t%s\t%s\t\t%s\t%s\t%s\t%s\t%s\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t
"""
splitTemplate = """\
SPL\t\t%s\t%s\t%s\t%s\t%s\t\t%s\t%s\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t
"""

# keep the same filename, just swap the extension
dir, file = os.path.split(file)
file, ext = file.split('.')
newFile = os.path.join(dir, file + '.iif')

typesNotImplemented = [set(), []]

with open(newFile, 'w') as f:
    f.write(fileStart)	

    tranFailures = []
    outgoing = ['Check']
    incoming = ['Deposit']

    for key, value in transactionMap.iteritems():

        # find out what the main transaction item is and what the split items are
        tType = value[0]['Type']
        if tType in outgoing:
            tranList = [ val for val in value if Decimal(val['Amount']) < 0 ]
            splList =  [ val for val in value if Decimal(val['Amount']) > 0 ]
        elif tType in incoming:
            tranList = [ val for val in value if Decimal(val['Amount']) > 0 ]
            splList =  [ val for val in value if Decimal(val['Amount']) < 0 ]
        else:
            typesNotImplemented[0].add(tType)
            typesNotImplemented[1].append(key)
            #tranFailures.append((value, "Transaction type not implemented"))
            continue

        if len(tranList) != 1:
            tranFailures.append((value, "There was a problem deciphering the data, possibly because the transaction amount is $0.00"))
            continue

        if len(tranList) > 1:
            print key, ' bigger than 1'

        # there is only one main transaction
        trans = tranList[0]
                
        # try to create a split if possible
        if len(splList) == 0:
            if trans['Split'] == '-SPLIT-':
                tranFailures.append((value, "Can not resolve the split values"))
            elif trans['Split'] == '':
                pass
            else:
                spl = dict(trans)
                spl['Amount'] = -1
                spl['AccountName'], spl['Split'] = spl['Split'], spl['AccountName']
                splList.append(spl)

        # test to make sure the splits add up
        tranSum = -Decimal(trans['Amount'])
        splSum = sum([ Decimal(val['Amount']) for val in splList ])
        if not tranSum == splSum:
            tranFailures.append((value, "The sum of the splits does not equal the total of the transaction"))
            continue

        # create the main transaction
        accountName = trans['AccountName'][-1]
        tType = trans['Type']
        date = trans['Date']
        name = trans['Name']
        amount = trans['Amount']
        memo = trans['Memo']
        cleared = 'N'
        toPrint = 'N'
        address1 = ''
        address2 = ''

        trans = transTemplate % (tType, date, accountName, name, amount, memo, cleared, toPrint, address1, address2)
        f.write(trans)

        # create the splits
        for spl in splList:
            splAccount = spl['AccountName'][-1]
            splAmount = spl['Amount']
            splCleared = 'N'
            splString = splitTemplate % (tType, date, splAccount, name, splAmount, memo, splCleared)
            f.write(splString)

        f.write(transEnd)

# handle failures
for failure in tranFailures:
    lbox.insert(Tkinter.END, '%s - %s' % (failure[0][0]['Trans #'], failure[1]))
if typesNotImplemented[0]:
    lbox.insert(Tkinter.END, 'Any transaction with a type of %s was not exported.  These transactions were: %s' % ([x for x in typesNotImplemented[0]], typesNotImplemented[1]))

lbox.insert(Tkinter.END, "File created: %s" % newFile)    
lbox.insert(Tkinter.END, "Finished processing file %s" % file)    
Tkinter.mainloop()

#!/usr/bin/python

from __future__ import with_statement
import sys, collections, operator, Tkinter, tkFileDialog

file = None

# find the file to import
try:
    file = sys.argv[1]
except IndexError:
    root = Tkinter.Tk()
    root.withdraw()
    file = tkFileDialog.askopenfilename(filetypes=[("Quickbooks Files", ".iif")])
    root.deiconify()

Tkinter.Label(text="Messages:").pack()
lbox = Tkinter.Listbox()
lbox.pack(expand=True, fill=Tkinter.BOTH, side=Tkinter.TOP)

if not file:
    lbox.insert(Tkinter.END, "No file to convert, quitting ...")
    Tkinter.mainloop()
    exit()

# return csv values in a list
def splitLine(line):
    columns = line.rstrip('\n')
    columnArray = columns.split(',')
    return [ val.rstrip('"').lstrip('"') for val in columnArray ]

transactionMap = collections.defaultdict(list)
with open(file) as f:

    # parse header columns
    columnArray = splitLine(f.readline())
    columnArray[0] = 'AccountName'
    
    currentAccount = []

    # loop through data
    for line in f:
        columns = splitLine(line)
        
        # if the first column has data it is the account data
        if columns[0]:
            if columns[0].startswith('Total'):
                currentAccount.pop()
            elif columns[0] == 'TOTAL':
                break;
            else:
               currentAccount.append(columns[0])
        else:
            # create a map of the data:  column->data
            newMap = {}
            for i, data in enumerate(columns):
                newMap[columnArray[i]] = data

            newMap['AccountName'] = list(currentAccount)

            transactionMap[newMap['Trans #']].append(newMap)
    
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

with open('test', 'w') as f:
    f.write(fileStart)	

    tranFailures = []
    for key, value in transactionMap.iteritems():
        
        # find out what the main transaction item is and what the split items are
        tType = value[0]['Type']
        if tType == "Check":
            tranList = [ val for val in value if float(val['Amount']) < 0 ]
            splList =  [ val for val in value if float(val['Amount']) > 0 ]
        elif tType == "Deposit":
            tranList = [ val for val in value if float(val['Amount']) > 0 ]
            splList =  [ val for val in value if float(val['Amount']) < 0 ]
        else:
            tranFailures.append((value, "Transaction type not implemented"))
            continue

        if len(tranList) != 1:
            tranFailures.append((value, "Could not decipher transaction, possibly the amount is $0.00"))
            continue

        # there is only one main transaction
        trans = tranList[0]
                
        # try to create a split if possible
        if len(splList) == 0:
            if trans['Split'] == '-SPLIT-':
                tranFailures.append((value, "Can not resolve the split values"))
            elif trans['Split'] == '':
                pass
            else:
                spl = trans
                spl['Amount'] *= -1
                spl['AccountName'], spl['Split'] = spl['Split'], spl['AccountName']
                splList.append(spl)
                
            # test to make sure the splits add up
            tranSum = -float(trans['Amount']) 
            splSum = sum([ float(val['Amount']) for val in splList ])
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

for failure in tranFailures:
    lbox.insert(Tkinter.END, failure)

lbox.insert(Tkinter.END, "Finished processing file %s" % file)    
Tkinter.mainloop()

import unittest
from qbexport import *
from decimal import Decimal
import StringIO

class MessageWindow(object):
    def insert(self):
        pass
    def show(self):
        pass

class TestSequenceFunctions(unittest.TestCase):
    
    def setUp(self):
        pass

    def testGetFileToWrite(self):
        files = [ ('test.csv', 'test.iif'), ('how.csv', 'how.iif') ]
        for file, result in files:
            self.assertTrue(result == getFileToWrite(file))

    def testCheckColumns(self):
        columns = [ 'AccountName', 'Trans #', 'Type', 'Split', 'Date', 'Name', 'Memo']
        self.assertRaises(ColumnsInvalidError, checkColumns, columns)
        columns.append('Amount')
        checkColumns(columns)
        columns.remove('Amount')
        columns.append('Debit')
        self.assertRaises(ColumnsInvalidError, checkColumns, columns)
        columns.append('Credit')
        checkColumns(columns)
        columns.remove('Debit')
        self.assertRaises(ColumnsInvalidError, checkColumns, columns)

    def testGetDecimal(self):
        self.assertTrue(0 == getDecimal(''))
        self.assertTrue(2 == getDecimal('2'))
        self.assertTrue(Decimal('2.0') == getDecimal('2.0'))

    def testGetDataRow(self):
        data = [

            ("Sterling MM #17656",'','','','','','','','','','',""),
            ('',"1322","Deposit","12/30/2000",'',"from year end",'','',"Opening Balance",'209988.90',"",'209988.90','-209988.90'),
            ('',"75","Deposit","4/11/2001",'',"transfer",'','',"#3923 (Reed)","",'10000.00','199988.90','10000'),
            ('',"106","Deposit","5/1/2001",'',"transfer",'','',"#3923 (Reed)","",'55000.00','144988.90','55000'),
            ("Total #3923 (Reed)",'','','','','','','','','2585108.20','2601597.58','-16489.38',''),
            ("Sterling MM #17656",'','','','','','','','','','',""),
            ('',"1875","Check","1/10/2004","5718","US Cellular",'','',"Phone","",'157.04','-11912.35','157.04'),
            ('',"1876","Check","1/10/2004","5719","Intermountain West Insulation",'','',"Insulation","",'4502.03','-16414.38','4502.03'),
            ('',"1877","Check","1/10/2004","5720","Ace Sales and Service",'','',"Misc,","",'75.00','-16489.38','75'),
            ("Total #3923 (Reed)",'','','','','','','','','2585108.20','2601597.58','-16489.38')
            ]
        results = [
            None,
            {'AccountName':['Sterling MM #17656'],'Trans #':"1322",'Type':"Deposit",'Date':"12/30/2000",'Num':'','Name':"from year end",'Memo':'','Clr':'','Split':"Opening Balance",'Debit':'209988.90','Credit':'','Balance':'209988.90', 'Amount':Decimal('-209988.90')},
            {'AccountName':['Sterling MM #17656'],'Trans #':"75",'Type':"Deposit",'Date':"4/11/2001",'Num':'','Name':"transfer",'Memo':'','Clr':'','Split':"#3923 (Reed)",'Debit':'','Credit':'10000.00','Balance':'199988.90', 'Amount':Decimal('10000.00')},
            {'AccountName':['Sterling MM #17656'],'Trans #':"106",'Type':"Deposit",'Date':"5/1/2001",'Num':'','Name':"transfer",'Memo':'','Clr':'','Split':"#3923 (Reed)",'Debit':'','Credit':'55000.00','Balance':'144988.90', 'Amount':Decimal('55000.00')},
            None,
            None,
            {'AccountName':['Sterling MM #17656'],'Trans #':"1875",'Type':"Check",'Date':"1/10/2004",'Num':"5718",'Name':"US Cellular",'Memo':'','Clr':'','Split':"Phone",'Debit':"",'Credit':'157.04','Balance':'-11912.35', 'Amount':Decimal('157.04')},
            {'AccountName':['Sterling MM #17656'],'Trans #':"1876",'Type':"Check",'Date':"1/10/2004",'Num':"5719",'Name':"Intermountain West Insulation",'Memo':'','Clr':'','Split':"Insulation",'Debit':'','Credit':'4502.03','Balance':'-16414.38', 'Amount':Decimal('4502.03')},
            {'AccountName':['Sterling MM #17656'],'Trans #':"1877",'Type':"Check",'Date':"1/10/2004",'Num':"5720",'Name':"Ace Sales and Service",'Memo':'','Clr':'','Split':"Misc,",'Debit':'','Credit':'75.00','Balance':'-16489.38', 'Amount':Decimal('75.00')},
            None
            ]

        columns = ['AccountName',"Trans #","Type","Date","Num","Name","Memo","Clr","Split","Debit","Credit","Balance", 'Nothin']

        for line, value in zip(data, results):
            result = getDataRow(line, columns)
            
            if result == None:
                self.assertTrue(value == None)
                continue
                
            for i, v in result.iteritems():
                if i in ['AccountName', 'Trans #', 'Type', 'Split', 'Date', 'Name', 'Memo', 'Credit', 'Debit', 'Amount']:
                    self.assertTrue(value[i] == v)

        columns = ['AccountName',"Trans #","Type","Date","Num","Name","Memo","Clr","Split","Deb","Cdit","Balance", 'Amount']

        for line, value in zip(data, results):
            result = getDataRow(line, columns)
            
            if result == None:
                self.assertTrue(value == None)
                continue
                
            for i, v in result.iteritems():
                if i in ['AccountName', 'Trans #', 'Type', 'Split', 'Date', 'Name', 'Memo', 'Credit', 'Debit', 'Amount']:
                    self.assertTrue(value[i] == v)


    def testIndexById(self):
        data = [
            {'Trans #':'1', 'id':1},
            {'Trans #':'1', 'id':2},
            {'Trans #':'1', 'id':3},
            {'Trans #':'2', 'id':4}
            ]
        result = {
            '1':[{'Trans #':'1', 'id':1}, {'Trans #':'1', 'id':2}, {'Trans #':'1', 'id':3}],
            '2':[{'Trans #':'2', 'id':4}]
            }

        r = indexById(data)
        self.assertTrue(r == result)

    def testWriteTransaction(self):
        trans = {'AccountName':['Test'], 'Type':'Check', 'Date':'1/2/34', 'Name':'Test', 'Amount':'50', 'Memo':'None'}
        splits = [{'AccountName':['asdf'], 'Amount':'34.50'}, {'AccountName':['asdf'], 'Amount':'34.50'}]
        file = StringIO.StringIO()
        writeTransaction(trans, splits, file)

        testval = """\
TRNS\t\t%s\t%s\t%s\t%s\t%s\t\t%s\t%s\t%s\t%s\t%s\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t
SPL\t\t%s\t%s\t%s\t%s\t%s\t\t%s\t%s\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t
SPL\t\t%s\t%s\t%s\t%s\t%s\t\t%s\t%s\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t
ENDTRNS
""" % ('Check', '1/2/34', 'Test', 'Test', '50', 'None', 'N', 'N', '', '', 'Check', '1/2/34', 'asdf', 'Test', '34.50', 'None', 'N', 'Check', '1/2/34', 'asdf', 'Test', '34.50', 'None', 'N')

        self.assertTrue(testval == file.getvalue())
        file.close()

    def testDecipherTransactions(self):
        columns = ['AccountName',"Trans #","Type","Date","Num","Name","Memo","Clr","Split","Debit","Credit","Balance", 'Nothin']
        data = [
            ("Sterling MM #17656",'','','','','','','','','','',""),
            ('',"1322","Check","12/30/2000",'',"from year end",'','',"Opening Balance",'209988.90',"",'209988.90','-209988.90'),
            ('',"75","Deposit","4/11/2001",'',"transfer",'','',"#3923 (Reed)","",'10000.00','199988.90','10000'),
            ('',"106","Deposit","5/1/2001",'',"transfer",'','',"#3923 (Reed)","",'55000.00','144988.90','55000'),
            ('',"108","Deposit","5/1/2001",'',"transfer",'','',"split","",'55000.00','144988.90','55000'),
            ('',"109","Deposit","5/1/2001",'',"transfer",'','',"-SPLIT-","",'55000.00','144988.90','55000'),
            ("Total #3923 (Reed)",'','','','','','','','','2585108.20','2601597.58','-16489.38',''),
            ("Test data",'','','','','','','','','','',""),
            ('',"106","Deposit","1/10/2004","5718","US Cellular",'','',"Phone","55000.00",'','-11912.35','157.04'),
            ('',"75","Deposit","1/10/2004","5718","US Cellular",'','',"Phone","",'10000.00','-11912.35','157.04'),
            ('',"1322","Check","1/10/2004","5719","Intermountain West Insulation",'','',"Insulation","",'0.01','-16414.38','4502.03'),
            ('',"1322","Check","1/10/2004","5720","Ace Sales and Service",'','',"Misc,","",'209988.89','-16489.38','75'),
            ("Total #3923 (Reed)",'','','','','','','','','2585108.20','2601597.58','-16489.38')
            ]

        transactions = [ getDataRow(x,columns) for x in data if getDataRow(x, columns) != None ]
        transactionMap = indexById(transactions)

        trans, splits = decipherTransactions(transactionMap['1322'])
        self.assertTrue(trans == transactions[0])
        self.assertTrue(splits == [ transactions[7], transactions[8] ])
        self.assertRaises(ParseError, decipherTransactions, transactionMap['75'])
        trans, splits = decipherTransactions(transactionMap['106'])
        self.assertTrue(trans == transactions[2])            
        self.assertTrue(splits == [ transactions[5]])
        trans, splits = decipherTransactions(transactionMap['108'])
        self.assertTrue(trans == transactions[3])
        self.assertTrue(len(splits) == 1)
        self.assertTrue(splits[0]['Amount'] == Decimal('-55000'))
        self.assertTrue(splits[0]['Split'] == 'Sterling MM #17656')
        self.assertTrue(splits[0]['AccountName'] == ['split'])
        self.assertRaises(ParseError, decipherTransactions, transactionMap['109'])


if __name__ == '__main__':
    unittest.main()

import os
from datetime import datetime

import pandas as pd
import pyodbc

dbPath = os.path.join(os.getcwd(), "CMSdatabase.accdb")



queryCaseDetails = '''SELECT CaseDetails.[caseYear], CaseDetails.[casePFSA], CaseDetails.[caseFTM], CaseDetails.[CaseNosAddl],
                        CaseDetails.[NoOfParcels], CaseDetails.[AnalystName], CaseDetails.[ReviewerName]
                        FROM CaseDetails
                        WHERE (((CaseDetails.[caseFTM])=
                        '''

queryCaseDetailsForIdentifiers = '''SELECT CaseDetails.Batch, CaseDetails.caseYear, CaseDetails.casePFSA, 
                                    CaseDetails.caseFTM, CaseDetails.Addressee, Items.District, Parcel.ParcelNo
                                    FROM (CaseDetails INNER JOIN Parcel ON CaseDetails.[caseFTM] 
                                    = Parcel.[CaseNoFK]) INNER JOIN Items ON Parcel.[ID] = Items.[ParcelNoFK]
                                    '''
queryParcelsDetails = '''SELECT Parcel.CaseNoFK, Parcel.ParcelNo, Parcel.SubmissionDate, Parcel.SubmitterName, 
                        Parcel.Rank, Items.FIR, Items.FIRDate, Items.EVCaliber, 
                        Items.EVType, Items.EV, Items.ItemNo, Items.Quantity
                        FROM Parcel INNER JOIN Items ON Parcel.[ID] = Items.[ParcelNoFK]
                        WHERE (((Parcel.CaseNoFK)=
                        '''
queryCOC = '''SELECT COC.[caseFTMFK], COC.[frmGRLDate], COC.[ProcessingDate], COC.[ComparisonStartDate], 
                COC.[ComparisonCompDate], COC.[ReviewStartDate], COC.[ReviewEndDate], COC.[BalScanStartDate], 
                COC.[BalScanCompDate], COC.[toCPRDate]
                FROM COC
                WHERE (((COC.[caseFTMFK])='''


#TODO Need to change Connectable ENGINE to SQLAlchemy
class AccessFile():
    def __init__(self) -> None:
        self.openConnection()

    def openConnection(self):
        try:
            conn_str = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                f'DBQ={dbPath};'
                )

            self.cnxn = pyodbc.connect(conn_str)
            self.crsr = self.cnxn.cursor()
            print('Database Opened.')
        except:
            print("connection to database is not established.")

    def closeConnection(self):
        self.cnxn.close()

    def readQuery(self, Query):
        df = pd.read_sql_query(Query, self.cnxn)
        if(df.empty):
            print("Reading Query Failure. No data found against this case number.")
        else:
            print("Reading Query Success. Data found against this case number.")
            return df


class Tables():
    def __init__(self, ftmNo) -> None:
        self.ftmNo = ftmNo
        self.database = AccessFile()

    def getTableByFtmNo(self, queryToRead:str) -> pd.DataFrame:
        return self.database.readQuery(f"{queryToRead} {self.ftmNo}));")
        

class CaseDetailsDF(Tables):
    def __init__(self, ftmNo) -> None:
        super().__init__(ftmNo)
        self.caseDetailsDF = self.getTableByFtmNo(queryCaseDetails)

    def getCaseNoParts(self) -> tuple:
        return self.caseDetailsDF.iloc[0]['caseYear'], self.caseDetailsDF.iloc[0]['casePFSA'], self.caseDetailsDF.iloc[0]['caseFTM'] 

    def getValuefrmCaseDetails(self, columnName, indexNumber=0):
        return self.caseDetailsDF.iloc[indexNumber][columnName]


class CoCDF(Tables):
    def __init__(self, ftmNo) -> None:
        super().__init__(ftmNo)
        self.cocDF = self.getTableByFtmNo(queryCOC)
    
    def getCOCdate(self, whichTypeOfDate : str) -> datetime:
        return self.cocDF.iloc[0][whichTypeOfDate].to_pydatetime()


class ParcelsDF(Tables):
    def __init__(self, ftmNo) -> None:
        super().__init__(ftmNo)
        self.parcelsDF = self.getTableByFtmNo(queryParcelsDetails)

    # Filters and Slices dataframe
    def getFirearmsOrAmmoDF(self, typeOfItems:list) -> pd.DataFrame:
        df1 = self.parcelsDF[self.parcelsDF['EVType'].isin(typeOfItems)]
        return df1[['ParcelNo', 'EVCaliber', 'EVType', 'EV', 'ItemNo', 'Quantity']]

    def getNoOfParcels(self) -> int:
        return len(self.parcelsDF.index)

    def getValuefrmParcels(self, columnName, indexNumber):
        return self.caseDetailsDF.iloc[indexNumber][columnName]


class IdentifiersDF(Tables):
    def __init__(self, BatchDate, ftmNo="",) -> None:
        super().__init__(ftmNo)
        self.BatchDate = BatchDate
        self.identifiersDF = self.getTableByBatchDate(queryCaseDetailsForIdentifiers)

    def getTableByBatchDate(self, queryToRead:str) -> pd.DataFrame:
        return self.database.readQuery(
            f"{queryToRead} WHERE (((CaseDetails.Batch)=#{self.BatchDate}#))").drop_duplicates(subset=['caseFTM'], keep='last')

    def getValuefrmParcels(self, columnName, indexNumber):
        return self.caseDetailsDF.iloc[indexNumber][columnName]



if __name__ == "__main__":

    d = CaseDetailsDF(123456)
    print(d.caseDetailsDF)

    i = IdentifiersDF("01/03/2022")
    print(i.identifiersDF)
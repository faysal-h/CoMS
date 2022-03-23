import os
from datetime import date, datetime

import pandas as pd
import pyodbc

dbPath = os.path.join(os.getcwd(), "CMSdatabase.accdb")

ammo = ['bullet', 'metal piece', 'cartridge case']

queryCaseDetails = '''SELECT CaseDetails.[caseYear], CaseDetails.[casePFSA], CaseDetails.[caseFTM], CaseDetails.[CaseNosAddl],
                        CaseDetails.[NoOfParcels], CaseDetails.[AnalystName], CaseDetails.[ReviewerName], 
                        CaseDetails.[Balscanner], CaseDetails.[TeamMember]
                        FROM CaseDetails
                        WHERE (((CaseDetails.[caseFTM])=
                        '''

queryCaseDetailsForIdentifiersDate = '''SELECT CaseDetails.Batch, CaseDetails.caseYear, CaseDetails.casePFSA, 
                                    CaseDetails.caseFTM, CaseDetails.Addressee, CaseDetails.CaseNosAddl, 
                                    Items.FIR, Items.FIRDate, Items.PS, Items.District, CaseDetails.NoOfParcels
                                    FROM (CaseDetails INNER JOIN Parcel ON CaseDetails.[caseFTM] = 
                                    Parcel.[CaseNoFK]) INNER JOIN Items ON Parcel.[ID] = Items.[ParcelNoFK]
                                    '''

queryCaseDetailsForIdentifiersFtm = '''SELECT CaseDetails.Batch, CaseDetails.caseYear, CaseDetails.casePFSA, 
                                    CaseDetails.caseFTM, CaseDetails.Addressee, CaseDetails.CaseNosAddl, 
                                    Items.FIR, Items.FIRDate, Items.PS, Items.District, CaseDetails.NoOfParcels
                                    FROM (CaseDetails INNER JOIN Parcel ON CaseDetails.[caseFTM] = Parcel.[CaseNoFK]) INNER JOIN Items ON Parcel.[ID] = Items.[ParcelNoFK]
                                    WHERE (((CaseDetails.caseFTM)= 
                                    '''

queryParcelsDetails = '''SELECT Parcel.CaseNoFK, Parcel.ParcelNo, Parcel.SubmissionDate, Parcel.SubmitterName, 
                        Parcel.Rank, Items.FIR, Items.FIRDate, Items.EVCaliber, 
                        Items.EVType, Items.EV, Items.ItemNo, Items.Quantity, Items.Notes
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

    def getValuefrmCaseDetails(self, columnName, indexNumber=0) -> str:
        return self.caseDetailsDF.iloc[indexNumber][columnName]
        

class CoCDF(Tables):
    def __init__(self, ftmNo) -> None:
        super().__init__(ftmNo)
        self.cocDF = self.getTableByFtmNo(queryCOC)
    
    def getCOCdate(self, whichTypeOfDate) -> datetime:
        return self.cocDF.drop_duplicates(subset=['caseFTMFK'], 
                                    keep='last').iloc[0][whichTypeOfDate].to_pydatetime()


    def getCOCdateString(self, whichTypeOfDate : str) -> str:
        dateToReturn = self.getCOCdate( whichTypeOfDate)
        if(type(dateToReturn) == type(pd.NaT)):
            return ""
        else:
            return dateToReturn.strftime("%d-%m-%Y")
            

class ParcelsDF(Tables):
    def __init__(self, ftmNo) -> None:
        super().__init__(ftmNo)
        self.parcelsDF = self.getTableByFtmNo(queryParcelsDetails)

    # Filters and Slices dataframe
    def getFirearmsOrAmmoDF(self, typeOfItems:list) -> pd.DataFrame:
        df1 = self.parcelsDF[self.parcelsDF['EV'].isin(typeOfItems)]
        return df1[['ParcelNo', 'EVCaliber', 'EVType', 'EV', 'ItemNo', 'Quantity']]

    def getNoOfParcels(self) -> int:
        return len(self.parcelsDF.index)

    def getValuefrmParcels(self, columnName, indexNumber):
        return self.caseDetailsDF.iloc[indexNumber][columnName]

    def getAmmoItemNos(self):
        ammoItemsDF = self.parcelsDF[self.parcelsDF['EVType'].isin(['ammo'])]
        ammoItemsList = ammoItemsDF['ItemNo']
        return (', ').join(ammoItemsList)

    def getAllItemNos(self):
        items = self.parcelsDF['ItemNo'].values.tolist()
        return (', ').join(items)

    # Manipulate dataframe for case Details in processing sheet
    def getParcelsDetailsForProcessingSheet(self):
        return self.parcelsDF.drop(['CaseNoFK', 'SubmissionDate', 'SubmitterName','Rank', 'FIR', 'FIRDate' ],
                                             axis=1).sort_values('ParcelNo').values.tolist()

class IdentifiersDF(Tables):

    def __init__(self, BatchDate="", ftmNo="") -> None:
        super().__init__(ftmNo)
        self.BatchDate = BatchDate
        if not (self.BatchDate) == "":
            self.identifiersDF = self.getTableByBatchDate(queryCaseDetailsForIdentifiersDate)
        else:
            self.identifiersDF = self.getTableByFtmNo(queryCaseDetailsForIdentifiersFtm)

    def getTableByBatchDate(self, queryToRead:str) -> pd.DataFrame:
        # extracts a dataframe contain values for creating identifiers.
        x = self.database.readQuery(
            f"{queryToRead} WHERE (((CaseDetails.Batch)=#{self.BatchDate}#))").drop_duplicates(subset=['caseFTM'], keep='last')

        # converts FIR date to string format and replaces original column
        x['FIRDate'] = x['FIRDate'].apply(lambda x: x.date().strftime('%d-%m-%Y')).values.tolist()

        return x

    def getTableByFtmNo(self, queryToRead:str) -> pd.DataFrame:
        # extracts a dataframe contain values for creating identifiers.
        x = self.database.readQuery(
            f"{queryToRead} {self.ftmNo}));").drop_duplicates(subset=['caseFTM'], keep='last')

        # converts FIR date to string format and replaces original column
        x['FIRDate'] = x['FIRDate'].apply(lambda x: x.date().strftime('%d-%m-%Y')).values.tolist()

        return x

    def getValuefrmIdentifiers(self, columnName, indexNumber):
        return self.caseDetailsDF.iloc[indexNumber][columnName]


if __name__ == "__main__":

    # d = CaseDetailsDF(123456)
    # print(d.getValuefrmCaseDetails('TeamMember'))


    # i = IdentifiersDF(ftmNo=123456)
    # print(i.identifiersDF.dtypes)
    # x = i.identifiersDF.drop(labels=['Batch'], axis=1)
    # print(x)
    # # f = i.getFirDateByBatchDate()
    # # print(i.combineCaseDetailsWithFIRDate())

    import inflect
    ie = inflect.engine()

    i = ParcelsDF(123456)
    x = i.getParcelsDetailsForProcessingSheet()
    print(x)
    
    
    parcelsDetailText = []
    oldParcel = 0
    
    for element, item in enumerate(x, start=1):

        quantityText = ie.number_to_words(item[5])
        parcelsDetailText.append(f"""Parcel  + {str(item[0])} : {quantityText} {str(item[1])} 
                                                caliber {item[3]} (Item {item[4]}) {item[6]}""")

    print(parcelsDetailText[0])
    # c = CoCDF(123456)e
    # print(type(c.getCOCdate("BalScanStartDate")))
    # print(type(c.getCOCdate("frmGRLDate")))
    # print(c.getCOCdate("BalScanStartDate"))
    # print(c.getCOCdate("frmGRLDate"))
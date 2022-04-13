import os
import logging
import urllib
from datetime import datetime
from dateutil.parser import parse

from sqlalchemy import create_engine
import sqlalchemy_access as sa_a
import sqlalchemy_access.pyodbc as sa_a_pyodbc

import pandas as pd

logging.basicConfig(level=logging.INFO)

dbPath = os.path.join(os.getcwd(), "CMSdatabase.accdb")

ammo = ['bullet', 'metal piece', 'cartridge case']

customDateFormat = "%d.%m.%Y"

queryCaseDetails = '''SELECT CaseDetails.[caseYear], CaseDetails.[casePFSA], CaseDetails.[caseFTM], CaseDetails.[CaseNosAddl],
                        CaseDetails.[NoOfParcels], CaseDetails.[AnalystName], CaseDetails.[ReviewerName], CaseDetails.[TestsRequest], 
                        CaseDetails.[Balscanner], CaseDetails.[TeamMember], CaseDetails.[Addressee], CaseDetails.[Batch]
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
                        Items.EVType, Items.EV, Items.ItemNo, Items.Quantity, Items.Notes, Items.PS, 
                        Items.District, Items.Accused
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
            connection_string = (
                                r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
                                f"DBQ={dbPath};"
                                r"ExtendedAnsiSQL=1;"
                                )
            connection_uri = f"access+pyodbc:///?odbc_connect={urllib.parse.quote_plus(connection_string)}"
            self.engine = create_engine(connection_uri)
        
            logging.info('Connection to Database Established.')
        
        except ValueError as e:
            logging.error(f"connection to database is not established.\n Error is : {e}")

    # def closeConnection(self):
    #     self.cnxn.close()

    def readQuery(self, Query):
        df = pd.read_sql_query(Query, self.engine)
        if(df.empty):
            logging.error("Reading Query Failure. No data found against this case number.")
            return df
        else:
            logging.info("Reading Query Success. Data found against this case number.")
            return df




class DataFrames():
    '''This class and its child classess read queries and manipulate data in the form of 
        PANDAS DATAFRAMES'''
    def __init__(self, ftmNo) -> None:
        self.ftmNo = ftmNo
        self.database = AccessFile()

    def getTableByFtmNo(self, queryToRead:str) -> pd.DataFrame:
        return self.database.readQuery(f"{queryToRead} {self.ftmNo}));")
        
    def checkIfCaseExist(self) -> bool:
        tempDF = self.database.readQuery(f"{queryCaseDetails} {self.ftmNo}));")
        if(tempDF.empty):
            return False
        else:
            return True

    def checkIfBatcDateExist(self, BatchDate) -> bool:
        tempDF = self.database.readQuery(
                    f"{queryCaseDetailsForIdentifiersDate} WHERE (((CaseDetails.Batch)=#{BatchDate}#))")
        if(tempDF.empty):
            return False
        else:
            return True

class CaseDetailsDF(DataFrames):
    '''class for manipulating DATAFRMAE of  Case Details Table in ACCESS DATABASE'''

    def __init__(self, ftmNo) -> None:
        super().__init__(ftmNo)
        self.caseDetailsDF = self.getTableByFtmNo(queryCaseDetails)

    def getCaseNoParts(self) -> tuple:
        return self.caseDetailsDF.iloc[0]['caseYear'], self.caseDetailsDF.iloc[0]['casePFSA'], self.caseDetailsDF.iloc[0]['caseFTM'] 

    def getValuefrmCaseDetails(self, columnName, indexNumber=0) -> str:
        return self.caseDetailsDF.iloc[indexNumber][columnName]
    
    def getBatchDate(self) -> datetime:
        return self.caseDetailsDF.iloc[0]['Batch'].to_pydatetime()
        

class CoCDF(DataFrames):
    def __init__(self, ftmNo) -> None:
        super().__init__(ftmNo)

        self.cocDF = self.getTableByFtmNo(queryCOC).drop_duplicates(
                                subset=['caseFTMFK'], keep='last')
    
    def getCOCdate(self, whichTypeOfDate) -> datetime:

        if(self.cocDF.empty):
            return ""
        else:
            x = self.cocDF.iloc[0][whichTypeOfDate]

            if(pd.isnull(x)):
                return ""
            else:
                return self.cocDF.iloc[0][whichTypeOfDate].to_pydatetime()


    def getCOCdateString(self, whichTypeOfDate : str) -> str:
        dateToReturn = self.getCOCdate( whichTypeOfDate)
        if((type(dateToReturn) == type(pd.NaT)) or dateToReturn == ""):
            return ""
        else:
            return dateToReturn.strftime(customDateFormat)
            

class ParcelsDF(DataFrames):
    def __init__(self, ftmNo) -> None:
        super().__init__(ftmNo)
        self.parcelsDF = (self.getTableByFtmNo(queryParcelsDetails)).sort_values(by=['ParcelNo'])

    # Filters and Slices dataframe
    def getFirearmsOrAmmoDF(self, typeOfItems:list) -> pd.DataFrame:
        df1 = self.parcelsDF[self.parcelsDF['EV'].isin(typeOfItems)]
        return df1[['ParcelNo', 'EVCaliber', 'EVType', 'EV', 'ItemNo', 'Quantity']]

    def getNoOfParcels(self) -> int:
        return len(self.parcelsDF.drop_duplicates(['ParcelNo']).index)

    def getValuefrmParcels(self, columnName, indexNumber):
        return self.parcelsDF.iloc[indexNumber][columnName]

    def getAmmoItemNos(self):
        ammoItemsDF = self.parcelsDF[self.parcelsDF['EVType'].isin(['ammo'])]
        ammoItemsList = ammoItemsDF['ItemNo']
        return (', ').join(ammoItemsList)

    def getAllItemNos(self):
        items = self.parcelsDF['ItemNo'].values.tolist()
        return (', ').join(items)

    def getDistrict(self):
        return self.parcelsDF.sort_values('SubmissionDate').drop_duplicates(subset=['CaseNoFK'], keep='last')

    # Manipulate dataframe for case Details in processing sheet
    def getParcelsDetailsForProcessingSheet(self):
        return self.parcelsDF.drop(['CaseNoFK', 'SubmissionDate', 'SubmitterName','Rank', 'FIR', 'FIRDate' ],
                                             axis=1).sort_values('ParcelNo').values.tolist()

    def getParcelDetailsForReport(self):
        parcelsForReport = self.parcelsDF.drop(['CaseNoFK'], axis=1).sort_values('ParcelNo')
        parcelsForReport['FIRDate'] = parcelsForReport['FIRDate'].apply(lambda x: x.date().strftime(customDateFormat)).values.tolist()
        parcelsForReport['SubmissionDate'] = parcelsForReport['SubmissionDate'].apply(lambda x: x.date().strftime(customDateFormat)).values.tolist()
        return parcelsForReport.values.tolist()

class IdentifiersDF(DataFrames):

    def __init__(self, BatchDate, ftmNo="") -> None:
        super().__init__(ftmNo)
        self.BatchDate = BatchDate
        if (self.BatchDate) != parse(BatchDate, fuzzy=False, dayfirst=True):
            self.identifiersDF = self.getTableByBatchDate(queryCaseDetailsForIdentifiersDate)
        else:
            self.identifiersDF = self.getTableByFtmNo(queryCaseDetailsForIdentifiersFtm)

    def getTableByBatchDate(self, queryToRead:str) -> pd.DataFrame:
        # extracts a dataframe contain values for creating identifiers.
        x = self.database.readQuery(
            f"{queryToRead} WHERE (((CaseDetails.Batch)=#{self.BatchDate}#))").drop_duplicates(subset=['caseFTM'], keep='last')

        # converts FIR date to string format and replaces original column
        x['FIRDate'] = x['FIRDate'].apply(lambda x: x.date().strftime(customDateFormat)).values.tolist()

        return x

    def getTableByFtmNo(self, queryToRead:str) -> pd.DataFrame:
        # extracts a dataframe contain values for creating identifiers.
        x = self.database.readQuery(
            f"{queryToRead} {self.ftmNo}));").drop_duplicates(subset=['caseFTM'], keep='last')

        # converts FIR date to string format and replaces original column
        x['FIRDate'] = x['FIRDate'].apply(lambda x: x.date().strftime(customDateFormat)).values.tolist()

        return x

    def getValuefrmIdentifiers(self, columnName, indexNumber):
        return self.caseDetailsDF.iloc[indexNumber][columnName]


if __name__ == "__main__":

    # d = CaseDetailsDF(123456)
    # print(type(d.getBatchDate()))
    # print(d.getValuefrmCaseDetails('TeamMember'))

    p = ParcelsDF(108185)


    print(p.parcelsDF)

    # i = IdentifiersDF('01/03')
    # print(i.identifiersDF)
    # x = i.identifiersDF.drop(labels=['Batch'], axis=1)
    # print(x)
    # # f = i.getFirDateByBatchDate()
    # # print(i.combineCaseDetailsWithFIRDate())


    # c = CoCDF(121212)
    # print(c.cocDF)
    # print(c.cocDF.empty)
    # print(c.getCOCdateString('ComparisonCompDate'))
    # # print(c.getCOCdateString('BalScanStartDate'))
    # bsDate = c.getCOCdate("BalScanStartDate")
    # print('bsDate')
    # print(bsDate)
    # # print(isinstance(bsDate, datetime))
    # print(type(c.getCOCdate("BalScanStartDate")))
    # print(type(c.getCOCdate("frmGRLDate")))
    # print(c.getCOCdate("BalScanStartDate"))
    # print(c.getCOCdate("frmGRLDate"))
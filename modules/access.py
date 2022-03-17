from msilib.schema import tables
import os
import pandas as pd
import pyodbc

dbPath = os.path.join(os.getcwd(), "CMSdatabase.accdb")

queryFtmNumber =    '''SELECT CaseDetails.caseYear, CaseDetails.casePFSA, CaseDetails.caseFTM, CaseDetails.NoOfParcels, CaseDetails.FrmGRLdate,
                        Parcel.ParcelNo, Parcel.SubmissionDate, Parcel.SubmitterName, Parcel.Rank, 
                        Items.FIR, Items.FIRDate, Items.EVCaliber, Items.EVType, Items.EV, Items.ItemNo, Items.Quantity
                        FROM (CaseDetails INNER JOIN Parcel ON CaseDetails.caseFTM = Parcel.CaseNoFK) INNER JOIN Items ON Parcel.ID = Items.ParcelNoFK;
                    '''

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
            print("No data found against this case number.")
        else:
            print("Data found against this case number.")
            return df



class Tables():
    def __init__(self, ftmNo) -> None:
        self.ftmNo = ftmNo
        self.database = AccessFile()
        self.cocDF = self.getCOCtable()

    def getCOCtable(self):

        queryCOC = f''' SELECT COC.[caseFTMFK], COC.[frmGRLDate], COC.[ProcessingDate], COC.[ComparisonStartDate], 
                        COC.[ComparisonCompDate], COC.[ReviewStartDate], COC.[ReviewEndDate], COC.[BalScanStartDate], 
                        COC.[BalScanCompDate], COC.[toCPRDate]
                        FROM COC
                        WHERE (((COC.[caseFTMFK])={self.ftmNo}));

                    '''
        return self.database.readQuery(queryCOC)



    def getCaseDetailsTable(self):
        queryCaseDetails = f''' SELECT CaseDetails.[caseYear], CaseDetails.[casePFSA], CaseDetails.[caseFTM], 
                                CaseDetails.[NoOfParcels], CaseDetails.[AnalystName], CaseDetails.[ReviewerName]
                                FROM CaseDetails
                                WHERE (((CaseDetails.[caseFTM])={self.ftmNo}));'''
        return self.database.readQuery(queryCaseDetails)

    def getParcelsTable(self):
        queryParcelsDetails = f''' SELECT Parcel.CaseNoFK, Parcel.ParcelNo, Parcel.SubmissionDate, Parcel.SubmitterName, 
                                Parcel.Rank, Items.FIR, Items.FIRDate, Items.EVCaliber, 
                                Items.EVType, Items.EV, Items.ItemNo, Items.Quantity
                                FROM Parcel INNER JOIN Items ON Parcel.[ID] = Items.[ParcelNoFK]
                                WHERE (((Parcel.CaseNoFK)={self.ftmNo}));
                                '''
        return self.database.readQuery(queryParcelsDetails)
        

    def getDate(self, whichDate):
        return self.cocDF.iloc[0][whichDate].to_pydatetime()


if __name__ == "__main__":


    db = Tables("123456")
    print(db.getCOCtable())
    print(db.getCaseDetailsTable())
    print(db.getParcelsTable())
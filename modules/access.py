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
        return pd.read_sql_query(Query, self.cnxn)
        

class DataFrames():
    pass


    


db = AccessFile()


df = db.readQuery(queryFtmNumber)
print(df) 
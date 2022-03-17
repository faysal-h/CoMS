import pandas as pd
import pyodbc



conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\Faisal\Documents\CMS\CMSdatabase.accdb;'
    )
cnxn = pyodbc.connect(conn_str)
crsr = cnxn.cursor()
for table_info in crsr.tables(tableType='TABLE'):
    print(table_info.table_name)

queryFtmNumber =    '''SELECT CaseDetails.caseYear, CaseDetails.casePFSA, CaseDetails.caseFTM, CaseDetails.NoOfParcels, CaseDetails.FrmGRLdate,
                        Parcel.ParcelNo, Parcel.SubmissionDate, Parcel.SubmitterName, Parcel.Rank, 
                        Items.FIR, Items.FIRDate, Items.EVCaliber, Items.EVType, Items.EV, Items.ItemNo, Items.Quantity
                        FROM (CaseDetails INNER JOIN Parcel ON CaseDetails.caseFTM = Parcel.CaseNoFK) INNER JOIN Items ON Parcel.ID = Items.ParcelNoFK;
                    '''

df = pd.read_sql_query(queryFtmNumber, cnxn)
cnxn.close()

 
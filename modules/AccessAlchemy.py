from operator import itemgetter
import os
import urllib
import pandas as pd
from sqlalchemy import create_engine

dbPath = os.path.join(os.getcwd(), "CMSdatabase.accdb")

connection_string = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    f"DBQ={dbPath};"
    r"ExtendedAnsiSQL=1;"
)
connection_uri = f"access+pyodbc:///?odbc_connect={urllib.parse.quote_plus(connection_string)}"
engine = create_engine(connection_uri)

queryCaseDetails = '''SELECT CaseDetails.[caseYear], CaseDetails.[casePFSA], CaseDetails.[caseFTM], CaseDetails.[CaseNosAddl],
                        CaseDetails.[NoOfParcels], CaseDetails.[AnalystName], CaseDetails.[ReviewerName], CaseDetails.[TestsRequest], 
                        CaseDetails.[Balscanner], CaseDetails.[TeamMember], CaseDetails.[Addressee]
                        FROM CaseDetails
                        '''

df = pd.read_sql_query(queryCaseDetails, engine)
print(df)

# Test if access driver engine installed
import pyodbc
[x for x in pyodbc.drivers() if x.startswith('Microsoft Access Driver')]


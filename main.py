#!/usr/bin/env python

import sys
import logging
from dateutil.parser import parse

import pymsgbox

from modules.AccessToDF import DataFrames
from modules import DocxEngine



logging.basicConfig(level=logging.DEBUG)

class menu():
    def __init__(self) -> None:
         self.choices = {
                        'Generate Sheets' : self.getCaseNoFromUser,
                        'Generate Identifiers' : self.getBatchDateFromUser,
                        "Quit": self.quitCMS
                        }

    def numericORlengthWarning(self):
        pymsgbox.alert(text="Enter a VALID 5 or 6 digits FTM number only.",
                            title="What are you doing?", button="Ok. I'm sorry")

    def wrongDateWarning(self):
        pymsgbox.alert(text="Enter valid date or Date does Not exist in Database.", title="Warning")
        self.run()

    def getCaseNoFromUser(self):
        ftmNo = pymsgbox.prompt(text='Enter FTM number', title="Sheet Generator")
        
        if(ftmNo is None):
            self.run
        else:
            if(self.validateCaseNumber(ftmNo)):
                self.generateSheets(ftmNo)
   
    def getBatchDateFromUser(self):
        batchDate = pymsgbox.prompt(text='Enter Batch Dat', title="Identifiers Generator")
        
        if(batchDate is None):
            self.run()
        else:
            parsedDate = self.parse_date(batchDate)
            if(DataFrames(ftmNo="").checkIfBatcDateExist(BatchDate=parsedDate)):
                self.generateIdentifiers(batchDate=batchDate)
            else:
                self.wrongDateWarning()
    
    def userPrompt(self):
        # return pymsgbox.prompt(text="Enter FTM Number for sheets.\nEnter BATCH DATE for Identifiers", title='CMS')
        return pymsgbox.confirm(text="What do you want to do?", title='CMS', 
                                buttons=[   'Generate Sheets', 'Generate Identifiers',
                                            'Quit'
                                        ]
                                )

    def run(self):
        while(True):
            choice = self.userPrompt()
            action = self.choices.get(choice)
            if action:
                action()
            else:
                pymsgbox.alert(text="Not a valid choice", title='Warning')

    def quitCMS(self):
        sys.exit(0)
    
    def validateCaseNumber(self, ftmNumber:int):
        if(ftmNumber.isnumeric()):
            if(len(str(ftmNumber)) in [5, 6]):
                if(DataFrames(ftmNumber).checkIfCaseExist()):
                    return True
                else:
                    self.numericORlengthWarning()
            else:
                self.numericORlengthWarning()
        else:
            self.numericORlengthWarning()

    def parse_date(self, batchDate):
        """
        Return whether the string can be interpreted as a date.

        :param string: str, string to check for date
        :param fuzzy: bool, ignore unknown tokens in string if True
        """
        try: 
            batchDate = parse(batchDate, fuzzy=False, dayfirst=True)
            return batchDate.strftime('%d/%m/%Y')

        except ValueError as e:
            logging.info(e)
            return self.wrongDateWarning()

    def generateSheets(self, ftmNumber):
        DocxEngine.ProcessingSheetProcessor(ftmNumber=ftmNumber).proceesingSheetMaker()
        DocxEngine.FirearmsProcessor(ftmNumber=ftmNumber).firearmSheetMaker()
        DocxEngine.CartridgeProcessor(ftmNumber=ftmNumber).cartridgeSheetMaker()
        DocxEngine.BulletProcessor(ftmNumber=ftmNumber).bulletSheetMaker()
        DocxEngine.ReportProcessor(ftmNumber=ftmNumber).reportGenerator()
        # pymsgbox.alert(text=f"All sheets are generated", title="Success")

    def generateIdentifiers(self, batchDate):        
        DocxEngine.IdentifiersProcessor(batchDate).FileIdentifierMaker()
        DocxEngine.IdentifiersProcessor(batchDate).EnvelopsMaker()
        DocxEngine.CPRProcessor(batchDate).FileCPRMaker()

if __name__ == "__main__":
    menu().run()
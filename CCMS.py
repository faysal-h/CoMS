import os
import sys
import logging
from dateutil.parser import parse

import pymsgbox

from modules.AccessToDF import DataFrames
from modules import DocxEngine

# create logger
logger = logging.getLogger('CCMS')
logger.setLevel(logging.DEBUG)

# # create handlers and set level to debug
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)

# # create formatter
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                              datefmt='%d-%m-%Y %I:%M:%S %p')

# # add formatter to ch
ch.setFormatter(formatter)

# # add ch to logger
logger.addHandler(ch)



class menu():
    def __init__(self) -> None:
        self.choices = {
            'Generate Sheets': self.getCaseNoFromUser,
            'Generate Identifiers': self.generateIdentifiers,
            'Batch Sheet Generator': self.generateSheetsInBatch,
            'Quit': self.quitCMS
        }

    def numericORlengthWarning(self):
        pymsgbox.alert(text="Enter a VALID 5 or 6 digits FTM number only.",
                            title="What are you doing?", button="Ok. I'm sorry")

    def wrongDateWarning(self):
        pymsgbox.alert(
            text="Enter valid date or Date does Not exist in Database.", title="Warning")
        self.run()

    def getCaseNoFromUser(self):
        ftmNo = pymsgbox.prompt(text='Enter FTM number',
                                title="Sheet Generator")

        if(ftmNo is None):
            self.run
        else:
            if(self._validateCaseNumber(ftmNo)):
                self.generateSheets(ftmNo)

    def getBatchDateFromUser(self):
        batchDate = pymsgbox.prompt(
            text='Enter Batch Date', title="Identifiers Generator")

        if(batchDate is None):
            self.run()
        else:
            parsedDate = self._parse_date(batchDate)
            if(DataFrames(ftmNo="").checkIfBatcDateExist(BatchDate=parsedDate)):
                return parsedDate
            else:
                self.wrongDateWarning()

    def userPrompt(self):
        # return pymsgbox.prompt(text="Enter FTM Number for sheets.\nEnter BATCH DATE for Identifiers", title='CMS')
        return pymsgbox.confirm(text="What do you want to do?", title='CMS',
                                buttons=[
                                    'Generate Sheets',
                                    'Batch Sheet Generator',
                                    'Generate Identifiers',
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

    def _validateCaseNumber(self, ftmNumber: int):
        if( 
            ftmNumber.isnumeric() and 
            len(str(ftmNumber)) in [5, 6] and 
            DataFrames(ftmNumber).checkIfCaseExist() ):
            return True
        else:
            self.numericORlengthWarning()
 
    def _parse_date(self, batchDate):
        """
        Return whether the string can be interpreted as a date.

        :param string: str, string to check for date
        :param fuzzy: bool, ignore unknown tokens in string if True
        """
        try:
            batchDate = parse(batchDate, fuzzy=False, dayfirst=True)
            return batchDate

        except ValueError as e:
            logger.error(e)
            return self.wrongDateWarning()

    def generateSheets(self, ftmNumber, openFolder=True):
        DocxEngine.ProcessingSheetProcessor(
            ftmNumber=ftmNumber).proceesingSheetMaker()
        DocxEngine.FirearmsProcessor(ftmNumber=ftmNumber).firearmSheetMaker()
        DocxEngine.CartridgeProcessor(
            ftmNumber=ftmNumber).cartridgeSheetMaker()
        DocxEngine.BulletProcessor(ftmNumber=ftmNumber).bulletSheetMaker()
        folderPath = DocxEngine.ReportProcessor(
            ftmNumber=ftmNumber).reportGenerator()
        if (openFolder is True):
            os.system(f"start {folderPath}")
        # pymsgbox.alert(text=f"All sheets are generated", title="Success")
        return folderPath

    def generateIdentifiers(self):
        batchDate = self.getBatchDateFromUser()
        logger.info(f'Batch date is {batchDate}')
        DocxEngine.IdentifiersProcessor(batchDate).FileIdentifierMaker()
        DocxEngine.IdentifiersProcessor(batchDate).EnvelopsMaker()
        DocxEngine.CPRProcessor(batchDate).FileCPRMaker()
        DocxEngine.NotesProcessor(batchDate).FileNotesMaker()

    def generateSheetsInBatch(self):
        batchDate = self.getBatchDateFromUser()
        cases = DocxEngine.IdentifiersProcessor(
            batchDate).getCasesInBatchDate()
        folderPath = ''
        for case in cases:
            try:
                folderPath = self.generateSheets(
                    ftmNumber=case, openFolder=False)
            except Exception as e:
                pymsgbox.alert(text=f'Data of Case {case} is not complete in Database',
                               title='Warning')
                logger.error(e)

        os.system(f"start {folderPath}")
        pymsgbox.alert(
            text=f"Sheets of all cases are generated", title="Success")

def main():
    logger.info('Starting')
    menu().run()
    logger.info('Ending')

if __name__ == "__main__":
    main()

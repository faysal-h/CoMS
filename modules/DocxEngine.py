import os
import re
import logging

import inflect
from docxtpl import DocxTemplate

from modules.CusPath import UserPaths

from modules.AccessToDF import CaseDetailsDF, CoCDF, ParcelsDF
from modules.AccessToDF import IdentifiersDF

from modules.identifierDocx import IdentifiersDocument, NotesDocument
from modules.CPRDocx import CPRDocument
from modules.reportDocx import Report

logger = logging.getLogger('CCMS.DocxEngine')


DateFormat = "%d.%m.%Y"

firearms = ['rifle', 'pistol', 'shotgun', 'machine pistol']
cartridge = ['cartridge case', 'cartridge cases', 'cartridge',
             'shotshell case', 'shotshell cases', 'shotshell']
bullet = ['bullet', 'metal piece', 'bullets', 'metal pieces', 'metallic piece', 'metallic pieces']

# List of caliber for which word "caliber" should be omitted.
SPECIAL_CALIBERS = ['12G', '9mm']

class IdentifiersProcessor():
    def __init__(self, batchDate) -> None:
        self.batchDate = batchDate

        # List of Identifiers from dataframe
        self.Identifiers = IdentifiersDF(
            self.batchDate).identifiersDF.values.tolist()

        # Creates Folder for each cases in identifiers
        self.currentBatchFolderPath = self.makeFoldersOfAllCasesInBatch()

        # BatchDate in string format
        self.fileNameEnder = self.batchDateToString()

    def zeroBeforFtmNumber(self, ftmNumber) -> str:
        if ftmNumber < 100000:
            return str(0) + str(ftmNumber)
        else:
            return str(ftmNumber)

    def batchDateToString(self) -> str:
        batchDateString = self.batchDate.strftime('%d-%m-%Y')
        return ''.join(ch for ch in batchDateString if ch.isalnum())

    def noneToEmptyValue(self, value):
        if(value == None):
            return ""
        else:
            return value

    def makeFoldersOfAllCasesInBatch(self):
        for identifier in self.Identifiers:

            caseNoFull = "PFSA" + str(identifier[1]) + "-" + str(
                identifier[2]) + "-FTM-" + self.zeroBeforFtmNumber(identifier[3])

            # Generates Respective case folder
            batchDate = identifier[0].to_pydatetime()
            # Batch Date Folder
            batchFolder = UserPaths().makeFolderfrmDate(date=batchDate)

            UserPaths().makeFolderInPath(batchFolder, caseNo=caseNoFull)

        return batchFolder

    def FileIdentifierMaker(self):
        i = IdentifiersDocument()
        i.PageLayout('A4')
        i.add_styles()
        i.createTwoColumnsPage()

        # i.addHeader(HeaderText=str("Batch Date: " + self.batchDate))

        for identifier in self.Identifiers:

            caseNoFull = "PFSA" + str(identifier[1]) + "-" + str(
                identifier[2]) + "-FTM-" + self.zeroBeforFtmNumber(identifier[3])
            caseNo2 = self.noneToEmptyValue(identifier[5])

            # generates identifiers for each case
            i.addFileIdentifiers(caseNo1=caseNoFull, caseNo2=str(caseNo2), parcels=str(identifier[10]),
                                 fir=str(identifier[6]), ps=str(identifier[8]),
                                 district=str(identifier[9]), BatchDate=self.batchDate.strftime('%d-%m-%Y'))

        i.saveDoc(os.path.join(self.currentBatchFolderPath,
                  f"Identifiers-{self.fileNameEnder}.docx"))

    def EnvelopsMaker(self):
        i = IdentifiersDocument()
        i.PageLayout('A4')
        i.add_styles()
        i.createTwoColumnsPage()

        for envelop in self.Identifiers:
            logger.info(envelop)

            caseNoFull = "PFSA" + \
                str(envelop[1]) + "-" + str(envelop[2]) + \
                "-FTM-" + self.zeroBeforFtmNumber(envelop[3])

            # i.tableIdentifiersFiles("PFSA2020-123456-FTM-123456", "PFSA2020-123456-FTM-123456", 1, "123 (XX.XX.XXXX)", "ABC&XYZ")
            i.addEnvelopsIdentifiers(
                caseNo1=caseNoFull, AddressTo=envelop[4], district=str(envelop[9]))

        i.saveDoc(os.path.join(self.currentBatchFolderPath,
                  f"Envelops-{self.fileNameEnder}.docx"))
        os.system(f"start {self.currentBatchFolderPath}")

    def getCasesInBatchDate(self):
        return [identifier[3] for identifier in self.Identifiers]


class NotesProcessor(IdentifiersProcessor):
    def __init__(self, batchDate) -> None:
        super().__init__(batchDate)


    def FileNotesMaker(self):
        # Uses Identifier Template for Creating Notes Sheet
        i = NotesDocument()
        i.PageLayout('A4')
        i.add_styles()
        # i.createTwoColumnsPage()

        for identifier in self.Identifiers:

            caseNoFull = "PFSA" + str(identifier[1]) + "-" + str(
                identifier[2]) + "-FTM-" + self.zeroBeforFtmNumber(identifier[3])
            caseNo2 = self.noneToEmptyValue(identifier[5])

            

            # Generates Note for each case
            i.addNote(caseNo1=caseNoFull, caseNo2=str(caseNo2), parcels=str(identifier[10]))

            # FTM no is identifier[3]
            p = ParcelsDF(int(identifier[3]))
            parcels = p.getParcelsDetailsForNotesSheet()
            # Add Parcels to Case
            for parcel in parcels:

                i.addParcelDetailsInNotes(parcelNo=parcel[0], itemNo=parcel[5], quantity=parcel[6],
                                         caliber=parcel[2], itemDetail= parcel[4])


        i.saveDoc(os.path.join(self.currentBatchFolderPath,
                f"Notes-{self.fileNameEnder}.docx"))


class CPRProcessor(IdentifiersProcessor):
    def __init__(self, batchDate) -> None:
        super().__init__(batchDate)

    def FileCPRMaker(self):
        cpr = CPRDocument()

        for i, identifier in enumerate(self.Identifiers, start=1):

            caseNoFull = "PFSA" + str(identifier[1]) + "-" + str(
                identifier[2]) + "-FTM-" + self.zeroBeforFtmNumber(identifier[3])

            # generates identifiers for each case
            cpr.addRowInMainTable(Serial=str(i), CaseNo=caseNoFull, FIR=str(identifier[6]),
                                  PS=str(identifier[8]), District=str(identifier[9]))

        cpr.save(os.path.join(self.currentBatchFolderPath,
                 f"CPR-{self.fileNameEnder}.docx"))


class Sheets():

    # Gets all required data in the form of dataframes and tables
    def __init__(self, ftmNumber) -> None:
        self.ftmNumber = ftmNumber

        # create instance of Dataframes to be used in all sheets processor
        self.caseDetailsDF = CaseDetailsDF(self.ftmNumber)
        self.CoCDF = CoCDF(self.ftmNumber)
        self.ParcelsDF = ParcelsDF(self.ftmNumber)

        # These variables will be used in all worksheets
        self.fullCaseNumber = self.fullCaseNumber()
        self.caseNumberParts = self.caseDetailsDF.getCaseNoParts()
        self.AdditionalCaseNumbers = self.secondCaseNoReplacer(
            self.caseDetailsDF.getValuefrmCaseDetails(columnName="CaseNosAddl"))
        self.analyst = self.caseDetailsDF.getValuefrmCaseDetails(
            columnName="AnalystName")
        self.reviewer = self.caseDetailsDF.getValuefrmCaseDetails(
            columnName="ReviewerName")
        self.addressee = self.caseDetailsDF.getValuefrmCaseDetails(
            columnName="Addressee")
        self.processingDate = self.CoCDF.getCOCdateString('ProcessingDate')
        self.BalscanDate = self.CoCDF.getCOCdateString("BalScanCompDate")
        self.toCPRdate = self.CoCDF.getCOCdateString('toCPRDate')

        # path of CASE Folder

        self.batchDate = self.caseDetailsDF.getBatchDate()
        self.batchFolderPath = UserPaths().makeFolderfrmDate(self.batchDate)
        self.currentCaseFolderPath = UserPaths.makeFolderInPath(
            self.batchFolderPath, self.fullCaseNumber)

    def secondCaseNoReplacer(self, caseNo2):
        '''This method replaces None case number with empty string'''
        if(caseNo2 in [None, "", "None"]):
            return ""
        else:
            return caseNo2

    def zeroBeforFtmNumber(self, ftmNumber) -> str:
        if ftmNumber < 100000:
            return str(0) + str(ftmNumber)
        else:
            return str(ftmNumber)

    def fullCaseNumber(self) -> str:
        x = self.caseDetailsDF.getCaseNoParts()
        return "PFSA" + str(x[0]) + "-" + str(x[1]) + "-FTM-" + self.zeroBeforFtmNumber(x[2])

    def numberToWord(self, digit):
        iE = inflect.engine()

        return iE.number_to_words(digit)


class ProcessingSheetProcessor(Sheets):
    # This classs manipulate data from DataFrames and varivale of COC, Case Details, Parcels
    def __init__(self, ftmNumber) -> None:
        super().__init__(ftmNumber)

        self.Balscanner = self.caseDetailsDF.getValuefrmCaseDetails(
            "Balscanner")
        self.TeamMember = self.caseDetailsDF.getValuefrmCaseDetails(
            "TeamMember")

        self.noOfParcels = self.ParcelsDF.getNoOfParcels()
        # Ammunition details in COC
        self.ammoItems = self._ammoItemsNoForCOC()
        
        # Total items in last cell of COC
        self.totalItemsNos = self._firearmItemsNoForCOC() + ', ' + self.ammoItems
        
        # Create instance of DOCX TEMPLATE for PROCESSING SHEET
        self.processingDocTemplate = DocxTemplate(UserPaths.processingTemplatePath)


    def __searchMinMaxNoInString(self, item:str) :
        if item not in ['', None]:
            result = [int(e) for e in re.split("[^0-9]", item) if e != '']
            logger.info(f'result in min and max is {result}')
            return min(result), max(result)
        else:
            return ''

    def __cocItemsStringMaker(self, searchedItems:list, itemString:list, itemLetter:str):
        #This method append items to items string list      
        if len(searchedItems) == 0:
            pass
        else:
            if searchedItems[0] == searchedItems[1]:
                itemString.append(f'{itemLetter}{min(searchedItems)}')
            else:
                itemString.append(f'{itemLetter}{min(searchedItems)} to {itemLetter}{max(searchedItems)}')

        return itemString

    def _ammoItemsNoForCOC(self) -> str:
        itemsList = self.ParcelsDF.getAmmoItemNos()
        logger.info(f'list of items are {itemsList}')
        # itemLetters = set(re.findall('\D', ''.join(self.itemsList)))

        # Ammo list to store items no
        cc = []
        ss = []
        mm = []
        bb = []

        # Separate different ammo items in two relevant list
        if len(itemsList) > 0:
            for i in itemsList:
                if i.capitalize().startswith('C'):
                    cc.append(i)
                elif i.capitalize().startswith('S'):
                    ss.append(i)
                elif i.capitalize().startswith('M'):
                    mm.append(i)
                elif i.capitalize().startswith('B'):
                    bb.append(i)
                else:
                    pass

        
        cc = self.__searchMinMaxNoInString(' '.join(cc))
        ss = self.__searchMinMaxNoInString(' '.join(ss))
        mm = self.__searchMinMaxNoInString(' '.join(mm))
        bb = self.__searchMinMaxNoInString(' '.join(bb))

        string = []

        self.__cocItemsStringMaker(cc, string, 'C')
        self.__cocItemsStringMaker(ss, string, 'SS')
        self.__cocItemsStringMaker(bb, string, 'B')
        self.__cocItemsStringMaker(mm, string, 'M')
        logger.info(f'Ammo Items string list is {string}')
        return ', '.join(string) + ', Test Fires'

    def _firearmItemsNoForCOC(self) -> str:
        itemsList = self.ParcelsDF.getFirearmsItemNos()
        logger.info(f'list of FIREARMS is {itemsList}')
        # Ammo list to store items no
        p = []
        s = []
        r = []
        m = []

        # Separate different ammo items in two relevant list
        if len(itemsList) > 0:
            for i in itemsList:
                if i.capitalize().startswith('P'):
                    p.append(i)
                elif i.capitalize().startswith('S'):
                    s.append(i)
                elif i.capitalize().startswith('R'):
                    r.append(i)
                elif i.capitalize().startswith('M'):
                    r.append(i)
                else:
                    pass

        p = self.__searchMinMaxNoInString(' '.join(p))
        s = self.__searchMinMaxNoInString(' '.join(s))
        r = self.__searchMinMaxNoInString(' '.join(r))
        m = self.__searchMinMaxNoInString(' '.join(m))
        logger.info(f'firearms {p}, {s}')
        string = []

        self.__cocItemsStringMaker(p, string, 'P')
        self.__cocItemsStringMaker(s, string, 'S')
        self.__cocItemsStringMaker(r, string, 'R')
        self.__cocItemsStringMaker(m, string, 'M')
        logger.info(f'Firearms String is {string}')
        return ', '.join(string)

    def _findTypeOfCOC(self):
        bs = self.Balscanner
        tm = self.TeamMember

        if(tm == None and bs == None):
            logger.info("single without balscan")
            return 1
        elif(tm == None and bs != None):
            logger.info('Single with balscan')
            return 2
        elif(bs == None):
            logger.info('team without balscan')
            return 3
        else:
            logger.info('team with balscan')
            return 4

    def setCoCandEVdetails(self):
        EvDetails = self.getAndSetParcels()
        x = self._findTypeOfCOC()
        if (x == 1):                                                # SINGLE WITHOUT BALSCAN
            return {
                'AGENCY_CASE': self.fullCaseNumber,
                'AGENCY_CASE2': self.AdditionalCaseNumbers,
                'ANALYST': self.analyst,
                'REVIEWER': self.reviewer,
                'PARCEL_DETAILS': EvDetails,
                'REVIEW_START': self.CoCDF.getCOCdateString('ReviewStartDate'),
                'REVIEW_END': self.CoCDF.getCOCdateString('ReviewEndDate'),
                'COMP_END': self.CoCDF.getCOCdateString('ComparisonCompDate'),

                'I1': self.noOfParcels,  # PARCELS COLUMN
                'I2': self.totalItemsNos,
                'I3': "",
                'I4': "",
                'I5': "",
                'I6': "",
                'A': self.analyst,                          # NAMES & INITIALS COLUMN
                'B': self.analyst,
                'C': "CPR",
                'D': "",
                'E': "",
                'F': "",
                'G': "",
                'H': "",
                'I': "",
                'J': "",
                'K': "",
                # DATE & TIME COLUMN
                'T1': self.CoCDF.getCOCdateString('frmGRLDate'),
                'T2': self.CoCDF.getCOCdateString('toCPRDate'),
                'T3': "",
                'T4': "",
                'T5': "",
                'T6': "",
                'P1': "CaseWork",            # PURPOSE COLUMN
                'P2': "Case Done",
                'P3': "",
                'P4': "",
                'P5': "",
                'P6': "",
            }
        elif(x == 2):                                                # SINGLE WITH BALSCAN
            return {
                'AGENCY_CASE': self.fullCaseNumber,
                'AGENCY_CASE2': self.AdditionalCaseNumbers,
                'ANALYST': self.analyst,
                'REVIEWER': self.reviewer,
                'PARCEL_DETAILS': EvDetails,
                'REVIEW_START': self.CoCDF.getCOCdateString('ReviewStartDate'),
                'REVIEW_END': self.CoCDF.getCOCdateString('ReviewEndDate'),
                'COMP_END': self.CoCDF.getCOCdateString('ComparisonCompDate'),

                'I1': self.noOfParcels,  # PARCELS COLUMN
                'I2': self.ammoItems,
                'I3': self.ammoItems,
                'I4': self.totalItemsNos,
                'I5': "",
                'I6': "",
                'A': self.analyst,                          # NAMES & INITIALS COLUMN
                'B': self.analyst,
                'C': self.Balscanner,
                'D': self.Balscanner,
                'E': self.analyst,
                'F': self.analyst,
                'G': "CPR",
                'H': "",
                'I': "",
                'J': "",
                'K': "",
                # DATE & TIME COLUMN
                'T1': self.CoCDF.getCOCdateString('frmGRLDate'),
                'T2': self.CoCDF.getCOCdateString('BalScanStartDate'),
                'T3': self.CoCDF.getCOCdateString('BalScanCompDate'),
                'T4': self.CoCDF.getCOCdateString('toCPRDate'),
                'T5': "",
                'T6': "",
                'P1': "CaseWork",            # PURPOSE COLUMN
                'P2': "BalScan",
                'P3': "BalScan Done",
                'P4': "Case Done",
                'P5': "",
                'P6': "",
            }
        elif(x == 3):                                     # TEAM WITH OUT BALSCAN
            return {
                'AGENCY_CASE': self.fullCaseNumber,
                'AGENCY_CASE2': self.AdditionalCaseNumbers,
                'ANALYST': self.analyst,
                'REVIEWER': self.reviewer,
                'PARCEL_DETAILS': EvDetails,
                'REVIEW_START': self.CoCDF.getCOCdateString('ReviewStartDate'),
                'REVIEW_END': self.CoCDF.getCOCdateString('ReviewEndDate'),
                'COMP_END': self.CoCDF.getCOCdateString('ComparisonCompDate'),

                'I1': self.noOfParcels,  # PARCELS COLUMN
                'I2': self.ammoItems,
                'I3': self.ammoItems,
                'I4': self.totalItemsNos,
                'I5': "",
                'I6': "",
                'A': self.TeamMember,                          # NAMES & INITIALS COLUMN
                'B': self.TeamMember,
                'C': self.analyst,
                'D': self.analyst,
                'E': self.TeamMember,
                'F': self.TeamMember,
                'G': "CPR",
                'H': "",
                'I': "",
                'J': "",
                'K': "",
                # DATE & TIME COLUMN
                'T1': self.CoCDF.getCOCdateString('frmGRLDate'),
                'T2': self.CoCDF.getCOCdateString('ComparisonStartDate'),
                'T3': self.CoCDF.getCOCdateString('ComparisonCompDate'),
                'T4': self.CoCDF.getCOCdateString('toCPRDate'),
                'T5': "",
                'T6': "",
                'P1': "CaseWork",            # PURPOSE COLUMN
                'P2': "Comparison",
                'P3': "Comparison Done",
                'P4': "Case Done",
                'P5': "",
                'P6': "",
            }
        else:                                           # SINGLE WITH BALSCAN
            return {
                'AGENCY_CASE': self.fullCaseNumber,
                'AGENCY_CASE2': self.AdditionalCaseNumbers,
                'ANALYST': self.analyst,
                'REVIEWER': self.reviewer,
                'PARCEL_DETAILS': EvDetails,
                'REVIEW_START': self.CoCDF.getCOCdateString('ReviewStartDate'),
                'REVIEW_END': self.CoCDF.getCOCdateString('ReviewEndDate'),
                'COMP_END': self.CoCDF.getCOCdateString('ComparisonCompDate'),


                'I1': self.noOfParcels,  # PARCELS COLUMN
                'I2': self.ammoItems,
                'I3': self.ammoItems,
                'I4': self.ammoItems,
                'I5': self.ammoItems,
                'I6': self.totalItemsNos,
                'A': self.TeamMember,                          # NAMES & INITIALS COLUMN
                'B': self.TeamMember,
                'C': self.analyst,
                'D': self.analyst,
                'E': self.Balscanner,
                'F': self.Balscanner,
                'G': self.analyst,
                'H': self.analyst,
                'I': self.TeamMember,
                'J': self.TeamMember,
                'K': "CPR",
                # DATE & TIME COLUMN
                'T1': self.CoCDF.getCOCdateString('frmGRLDate'),
                'T2': self.CoCDF.getCOCdateString('ComparisonStartDate'),
                'T3': self.CoCDF.getCOCdateString('ComparisonCompDate'),
                'T4': self.CoCDF.getCOCdateString('BalScanStartDate'),
                'T5': self.CoCDF.getCOCdateString('BalScanCompDate'),
                'T6': self.CoCDF.getCOCdateString('toCPRDate'),
                'P1': "CaseWork",            # PURPOSE COLUMN
                'P2': "Comparison",
                'P3': "Comparison Done",
                'P4': "Balscan",
                'P5': "Balscan Done",
                'P6': "Case Done",
            }

    def parcelDetailsStringMaker(self, itemDetails:list, sameParcel=False):
        ParcelNo = itemDetails[0]
        Quantity = int(itemDetails[5])
        EVCaliber = itemDetails[1]
        EVDetails = itemDetails[3]
        ItemsNo = itemDetails[4]
        notes = itemDetails[6]

        # Concatenate Items Details in a single String
        if notes in [None, "None"]:
            notes = ""
        # # converts digit to text
        # inflectEngine = inflect.engine()
        # Q = inflectEngine.number_to_words(Quantity)
        if Quantity < 2:
            itemWord = 'Item'
        else:
            itemWord = 'Items'

        caliberWord = ''
        if EVCaliber not in SPECIAL_CALIBERS:
            caliberWord = 'caliber '

        if(sameParcel == True):
            return (f"and {self.numberToWord(Quantity)} {str(EVCaliber)} {caliberWord}{EVDetails} ({itemWord} {ItemsNo}) {notes}")
        else:
            return (f"Parcel {str(ParcelNo)}: {self.numberToWord(Quantity)} {str(EVCaliber)} {caliberWord}{EVDetails} ({itemWord} {ItemsNo}) {notes}")

    def getAndSetParcels(self):
        # gets LIST of parcels in case and combine it to a single string.

        parcels = self.ParcelsDF.getParcelsDetailsForProcessingSheet()

        caseDetailsList = []

        for index, item in enumerate(parcels):
            logger.info(f"index {index}")
            # for parcel add PARCEL 1 to start
            if(index == 0):
                caseDetailsList.append(self.parcelDetailsStringMaker(itemDetails=item))
            else:
                # gets the parcel No of previous parcel
                oldParcelNo = parcels[index-1][0]

                if(item[0] != oldParcelNo):
                    caseDetailsList.append(self.parcelDetailsStringMaker(itemDetails=item))
                else:
                    caseDetailsList.append(self.parcelDetailsStringMaker(itemDetails=item, sameParcel=True))
        # this method also joins parcel string and returns a single string of case details for
        # processing sheet
        logger.info(caseDetailsList)
        return ("").join(caseDetailsList)

    def proceesingSheetMaker(self):
        # Poplulate and Generate Processing Sheet
        context = self.setCoCandEVdetails()
        self.processingDocTemplate.render(context)
        self.processingDocTemplate.save(os.path.join(self.currentCaseFolderPath,
                                        f'1. Processing Sheet-{self.ftmNumber}.docx'))


class FirearmsProcessor(Sheets):

    def __init__(self, ftmNumber) -> None:
        super().__init__(ftmNumber)

        # List of firearms
        self.firearms = self.ParcelsDF.getFirearmsOrAmmoDF(
            firearms).sort_values('ParcelNo').values.tolist()
        # Create instance of DOCX TEMPLATE
        self.firearmsDocTemplate = DocxTemplate(UserPaths.firearmsTemplatePath)

    def testFiresFromItemNo(self, itemNo: str) -> str:

        testFireLetter = 'C'
        if itemNo.upper().startswith('S'):
            testFireLetter = 'S'

        if itemNo in [None, "", " "]:
            return ""
        else:
            itemNo = itemNo.upper()
            return f'{itemNo}T{testFireLetter}1 & {itemNo}T{testFireLetter}2'
    # Iterate through each firearm in firarsm List and save a worksheet with corresponding item No

    def firearmSheetMaker(self):
        if len(self.firearms) > 0:
            for firearm in self.firearms:

                testFires = self.testFiresFromItemNo(firearm[4])
                yearShort = str(self.caseNumberParts[0])

                context = {
                    'AGENCY_CASE': self.fullCaseNumber,
                    'AGENCY_CASE2': self.AdditionalCaseNumbers,
                    'ITEM': firearm[4],
                    'EXAMINER': self.analyst,
                    'REVIEWER': self.reviewer,
                    'DATE': self.processingDate,
                    'CALIBER': firearm[1],
                    'FTMNO': self.caseNumberParts[2],
                    'MARKING': str(firearm[4])+"/"+self.zeroBeforFtmNumber(self.caseNumberParts[2])+"/"+yearShort[-2:],
                    'ABIS': self.BalscanDate,
                    'TESTFIRES': testFires,
                    'NOTES': firearm[6]
                }

                self.firearmsDocTemplate.render(context)
                self.firearmsDocTemplate.save(os.path.join(self.currentCaseFolderPath,
                                                           f"2. firearms-{self.ftmNumber}-{firearm[0]}.docx"))
        else:
            print("No firearms sheet is generated as no data is passed to processor")


class CartridgeProcessor(Sheets):

    def __init__(self, ftmNumber) -> None:
        super().__init__(ftmNumber)

        # List of firearms
        self.cartridges = self.ParcelsDF.getFirearmsOrAmmoDF(
            cartridge).sort_values('ParcelNo').values.tolist()
        # Create instance of DOCX TEMPLATE
        self.cartridgeTemplate = DocxTemplate(UserPaths.cartridgeTemplatePath)

    # Iterate through each firearm in firarsm List and save a worksheet with corresponding item No

    def cartridgeSheetMaker(self):
        context = {
            'AGENCY_CASE': self.fullCaseNumber,
            'AGENCY_CASE2': self.AdditionalCaseNumbers,
            'EXAMINER': self.analyst,
            'REVIEWER': self.reviewer,
            'DATE': self.processingDate,
            # 'CALIBER' : cartridge[1],
            # 'FTMNO' : self.caseNumberParts[2],
            # 'MARKING': str(cartridge[4])+"/"+str(self.caseNumberParts[1])+"/"+yearShort[2:],
            # 'ABIS': self.BalscanDate.strftime(DateFormat),
        }
        self.cartridgeTemplate.render(context)
        self.cartridgeTemplate.save(os.path.join(self.currentCaseFolderPath,
                                                 f"3. cartridge-{self.ftmNumber}.docx"))


class BulletProcessor(Sheets):

    def __init__(self, ftmNumber) -> None:
        super().__init__(ftmNumber)

        # List of firearms
        self.bullets = self.ParcelsDF.getFirearmsOrAmmoDF(
            bullet).sort_values('ParcelNo').values.tolist()
        # Create instance of DOCX TEMPLATE
        self.bulletDocTemplate = DocxTemplate(UserPaths.bulletTemplatePath)

    # Iterate through each firearm in firarsm List and save a worksheet with corresponding item No

    def bulletSheetMaker(self):
        if (len(self.bullets) > 0):
            for bullet in self.bullets:
                #     # yearShort = str(self.caseNumberParts[0])

                context = {
                    'AGENCY_CASE': self.fullCaseNumber,
                    'AGENCY_CASE2': self.AdditionalCaseNumbers,
                    'EXAMINER': self.analyst,
                    'REVIEWER': self.reviewer,
                    'DATE': self.processingDate,
                    # 'CALIBER' : cartridge[1],
                    # 'FTMNO' : self.caseNumberParts[2],
                    # 'MARKING': str(cartridge[4])+"/"+str(self.caseNumberParts[1])+"/"+yearShort[2:],
                    # 'ABIS': self.BalscanDate.strftime(DateFormat),
                }
                self.bulletDocTemplate.render(context)
                self.bulletDocTemplate.save(os.path.join(self.currentCaseFolderPath,
                                                         f"4. bullet-{self.ftmNumber}.docx"))


class ReportProcessor(Sheets):
    def __init__(self, ftmNumber) -> None:
        super().__init__(ftmNumber)

        self.numberOfParcels = self.ParcelsDF.getNoOfParcels()
        self.district = self.ParcelsDF.getValuefrmParcels(
            'District', indexNumber=0)
        # self.testStatement = self.caseDetailsDF.getValuefrmCaseDetails(columnName="TestsRequest")

        self.parcels = self.ParcelsDF.getParcelDetailsForReport()

        # it should come after self.parcels
        self.testStatement = self.testRequestHeading()

    def testRequestHeading(self):
        # As set cannot contains duplicate value
        cc = 0
        ss = 0
        firearm = 0

        # Finds and adds Cartridge case and shotshell case
        # Also sets Functionlity Testin True if firearm is present
        for parcel in self.parcels:
            if str(parcel[8]).lower().find('cart') != -1:
                cc = cc + parcel[10]
            elif str(parcel[8]).lower().find('shot') != -1:
                ss = ss + parcel[10]
            else:
                firearm = firearm + 1

        cartridgeCase = ''
        # Adds cartridge case in test request statement depending upon number of cartridges
        if cc > 0:
            if cc > 1:
                cartridgeCase = 'Cartridge Cases'
            else:
                cartridgeCase = 'Cartridge Case'

        shotShellCase = ''
        # Adds Shotshell cases in test request statement depending upon number of cartridges
        if ss > 0:
            if ss > 1:
                shotShellCase = 'Shotshell Cases'
            else:
                shotShellCase = 'Shotshell Case'

        f = ''
        if firearm > 1:
            f = 'Firearms'
        else:
            f = 'Firearm'

        And = ''
        if cc > 0 and ss > 0:
            And = ' and '

        funcTest = ''
        # Adds functionlity in Test request if firearm is present
        if firearm > 0:
            funcTest = ' and Functionality Testing'
        else:
            funcTest = ''

        return f"Comparison of {cartridgeCase}{And}{shotShellCase} with Submitted {f}{funcTest}"

    def reportGenerator(self):

        testReport = Report()
        testReport.PageLayout('A4')

        # testReport.add_styles()
        # testReport.paraTOD()

        testReport.tableCaseDetails(caseNo1=self.fullCaseNumber, caseNo2=self.AdditionalCaseNumbers,
                                    addressee=self.addressee, district=self.district)

        testReport.paraEvDetail(Addressee=self.addressee,
                                items=self.numberOfParcels,
                                District=self.district,
                                testRequest=self.testStatement)

        testReport.tableEvDetails(self.parcels)

        testReport.tableAnalysisDetails(
            startDate=self.processingDate, endDate=self.toCPRdate)

        # testReport.paraResults()
        # testReport.paraNotes()
        # testReport.paraDisposition()
        # testReport.footer()
        # adds header from second page onwards
        testReport.header(caseNo=self.fullCaseNumber)
        testReport.save(os.path.join(self.currentCaseFolderPath,
                        f'Report {self.ftmNumber}.docx'))

        return self.currentCaseFolderPath


if __name__ == "__main__":

    # r = ReportProcessor(123456)
    # print(r.parcels, end="\n")
    # print(r.testRequestHeading())

    # r.reportGenerator()

    n = NotesProcessor('08/03/2022')
    n.FileNotesMaker()

    # i = IdentifiersProcessor("1/3/2022")
    # # i.FileIdentifierMaker()
    # print(i.batchDate)
    # i.FileIdentifierMaker()
    # i.EnvelopsMaker()

    # s = Sheets(123456)
    # print(s.caseDetailsDF)
    # print(s.batchFolderPath)
    # print(s.currentCaseFolderPath)

    # i.FileIdentifierMaker()
    # i.EnvelopsMaker()

    p = ProcessingSheetProcessor(104307)
    print(p._ammoItemsNoForCOC())

    # p.proceesingSheetMaker(UserPaths.checkNcreateUserCaseWorkFolder())

    # f = FirearmsProcessor(123456)
    # print(f.currentCaseFolderPath)
    # print(f.firearmSheetMaker())

    # c = CartridgeProcessor(123456)
    # # print(c.cartridges)
    # c.cartridgeSheetMaker()

    # b = BulletProcessor(123456)
    # print(len(b.bullets))

    # print(UserPaths().checkNcreateCaseWorkDirectory())

    os.system(f"start {n.currentBatchFolderPath}")
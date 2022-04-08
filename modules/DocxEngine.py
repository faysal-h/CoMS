
import logging
import os

import inflect
from docxtpl import DocxTemplate

from modules.CusPath import UserPaths

from modules.AccessToDF import CaseDetailsDF, CoCDF, ParcelsDF
from modules.AccessToDF import IdentifiersDF

from modules.identifierDocx import IdentifiersDocument
from modules.CPRDocx import CPRDocument
from modules.reportDocx import Report

processingTemplatePath = os.path.join(os.getcwd(), "modules\\templates\\processing.docx")
firearmsTemplatePath = os.path.join(os.getcwd(), "modules\\templates\\firearms.docx")
cartridgeTemplatePath = os.path.join(os.getcwd(), "modules\\templates\\cartridge.docx")
bulletTemplatePath = os.path.join(os.getcwd(), "modules\\templates\\bullet.docx")

DateFormat = "%d.%m.%Y"

firearms = ['rifle', 'pistol', 'shotgun', 'machine pistol']
cartridge = ['cartridge case', 'cartridge cases','cartridge', 'shotshell case', 'shotshell cases', 'shotshell']
bullet = ['bullet', 'metal piece','bullets', 'metal pieces']


class IdentifiersProcessor():
    def __init__(self, batchDate) -> None:
        self.batchDate = batchDate

        # List of Identifiers from dataframe
        self.Identifiers = IdentifiersDF(self.batchDate).identifiersDF.values.tolist()

        # Creates Folder for each cases in identifiers
        self.currentBatchFolderPath = self.makeFoldersOfAllCasesInBatch()

        # BatchDate in string format
        self.fileNameEnder = self.batchDateToString()

    def batchDateToString(self) -> str:
        return ''.join(ch for ch in str(self.batchDate) if ch.isalnum())


    def noneToEmptyValue(self, value):
        if(value==None):
            return ""
        else:
            return value

    def makeFoldersOfAllCasesInBatch(self):
        for identifier in self.Identifiers:

            caseNoFull = "PFSA" + str(identifier[1]) + "-" + str(identifier[2]) + "-FTM-" + str(identifier[3]) 

            # Generates Respective case folder
            batchDate = identifier[0].to_pydatetime()
            batchFolder = UserPaths().makeFolderfrmDate(date=batchDate)
            UserPaths().makeFolderInPath(batchFolder, caseNo=caseNoFull)
        
        return batchFolder

    def FileIdentifierMaker(self):
        i = IdentifiersDocument()
        i.PageLayout('A4')
        i.add_styles()
        i.createTwoColumnsPage()

        for identifier in self.Identifiers:

            caseNoFull = "PFSA" + str(identifier[1]) + "-" + str(identifier[2]) + "-FTM-" + str(identifier[3]) 
            caseNo2 = self.noneToEmptyValue(identifier[5])

            # generates identifiers for each case
            i.addFileIdentifiers(caseNo1=caseNoFull, caseNo2=str(caseNo2), parcels=str(identifier[10]),
                                fir=str(identifier[6]), firDate=identifier[7], ps=str(identifier[8]),
                                district=str(identifier[9]))

        i.saveDoc(os.path.join(self.currentBatchFolderPath, f"Identifiers-{self.fileNameEnder}.docx"))

    def EnvelopsMaker(self):
        i = IdentifiersDocument()
        i.PageLayout('A4')
        i.add_styles()
        i.createTwoColumnsPage()

        for envelop in self.Identifiers:
            logging.info(envelop)

            caseNoFull = "PFSA" + str(envelop[1]) + "-" + str(envelop[2]) + "-FTM-" + str(envelop[3]) 

            # i.tableIdentifiersFiles("PFSA2020-123456-FTM-123456", "PFSA2020-123456-FTM-123456", 1, "123 (XX.XX.XXXX)", "ABC&XYZ")
            i.addEnvelopsIdentifiers(caseNo1=caseNoFull, AddressTo=envelop[4],district=str(envelop[9]) )

        i.saveDoc(os.path.join(self.currentBatchFolderPath, f"Envelops-{self.fileNameEnder}.docx"))
        os.system(f"start {self.currentBatchFolderPath}")

class CPRProcessor(IdentifiersProcessor):
    def __init__(self, batchDate) -> None:
        super().__init__(batchDate)

    def FileCPRMaker(self):
        cpr = CPRDocument()

        for i, identifier in enumerate(self.Identifiers, start=1):

            caseNoFull = "PFSA" + str(identifier[1]) + "-" + str(identifier[2]) + "-FTM-" + str(identifier[3]) 
            
            # generates identifiers for each case
            cpr.addRowInMainTable(Serial=str(i), CaseNo=caseNoFull, FIR=str(identifier[6]),
                                 FIRDate=identifier[7], PS=str(identifier[8]),
                                District=str(identifier[9]))

        cpr.save(os.path.join(self.currentBatchFolderPath, f"CPR-{self.fileNameEnder}.docx"))
    

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
        self.AdditionalCaseNumbers = self.secondCaseNoReplacer(self.caseDetailsDF.getValuefrmCaseDetails(columnName="CaseNosAddl"))
        self.analyst = self.caseDetailsDF.getValuefrmCaseDetails(columnName="AnalystName")
        self.reviewer = self.caseDetailsDF.getValuefrmCaseDetails(columnName="ReviewerName")
        self.addressee = self.caseDetailsDF.getValuefrmCaseDetails(columnName="Addressee")
        self.processingDate = self.CoCDF.getCOCdateString('ProcessingDate')
        self.BalscanDate = self.CoCDF.getCOCdateString("BalScanCompDate")
        self.toCPRdate = self.CoCDF.getCOCdateString('toCPRDate')

        #path of CASE Folder

        self.batchDate = self.caseDetailsDF.getBatchDate()
        self.batchFolderPath = UserPaths().makeFolderfrmDate(self.batchDate)
        self.currentCaseFolderPath = UserPaths.makeFolderInPath(self.batchFolderPath, self.fullCaseNumber)


    def secondCaseNoReplacer(self, caseNo2):
        '''This method replaces None case number with empty string'''
        if(caseNo2 in [None, "", "None"] ):
            return ""
        else:
            return caseNo2

    def fullCaseNumber(self) -> str:
        x = self.caseDetailsDF.getCaseNoParts()
        return "PFSA"+ str(x[0]) + "-" + str(x[1]) + "-FTM-" + str(x[2])

    def numberToWord(self, digit):
        iE = inflect.engine()

        return iE.number_to_words(digit)

#This classs manipulate data from DataFrames and varivale of COC, Case Details, Parcels
class ProcessingSheetProcessor(Sheets):
    def __init__(self, ftmNumber) -> None:
        super().__init__(ftmNumber)

        self.Balscanner = self.caseDetailsDF.getValuefrmCaseDetails("Balscanner")
        self.TeamMember = self.caseDetailsDF.getValuefrmCaseDetails("TeamMember")

        self.noOfParcels = self.ParcelsDF.getNoOfParcels()
        self.ammoItems = self.ParcelsDF.getAmmoItemNos() + ', Test Fires'
        self.totalItemsNos = self.ParcelsDF.getAllItemNos() + ', Test Fires'          
        # Create instance of DOCX TEMPLATE for PROCESSING SHEET
        self.processingDocTemplate = DocxTemplate(processingTemplatePath)

    def findTypeOfCOC(self):
        bs = self.Balscanner
        tm = self.TeamMember

        if(tm == None and bs == None):
            print("single withouth balscan")
            return 1
        elif(tm == None and bs != None):
            print('Single with balscan')
            return 2
        elif(bs == None):
            print('team without balscan')
            return 3
        else:
            print('team with balscan')
            return 4

    def setCoCandEVdetails(self):
        EvDetails = self.getAndSetParcels()
        x = self.findTypeOfCOC()
        if (x==1) :                                                # SINGLE WITHOUT BALSCAN
            return  {   
                        'AGENCY_CASE' : self.fullCaseNumber,
                        'AGENCY_CASE2' : self.AdditionalCaseNumbers,
                        'ANALYST': self.analyst,
                        'REVIEWER': self.reviewer,
                        'PARCEL_DETAILS': EvDetails,
                        'REVIEW_START': self.CoCDF.getCOCdateString('ReviewStartDate'),
                        'REVIEW_END': self.CoCDF.getCOCdateString('ReviewEndDate'),
                        'COMP_END': self.CoCDF.getCOCdateString('ComparisonCompDate'),
                        
                        'I1': self.noOfParcels,                #PARCELS COLUMN
                        'I2': self.totalItemsNos,
                        'I3': "",
                        'I4': "",
                        'I5': "",
                        'I6': "",
                        'A' : self.analyst,                          # NAMES & INITIALS COLUMN
                        'B' : self.analyst,
                        'C': "CPR",
                        'D': "",
                        'E': "",
                        'F' : "",
                        'G' : "",
                        'H' : "",
                        'I': "",
                        'J': "",
                        'K': "",
                        'T1': self.CoCDF.getCOCdateString('frmGRLDate'),           # DATE & TIME COLUMN
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
        elif(x==2):                                                # SINGLE WITH BALSCAN
            return  {   
                        'AGENCY_CASE' : self.fullCaseNumber,
                        'AGENCY_CASE2' : self.AdditionalCaseNumbers,
                        'ANALYST': self.analyst,
                        'REVIEWER': self.reviewer,
                        'PARCEL_DETAILS': EvDetails,
                        'REVIEW_START': self.CoCDF.getCOCdateString('ReviewStartDate'),
                        'REVIEW_END': self.CoCDF.getCOCdateString('ReviewEndDate'),
                        'COMP_END': self.CoCDF.getCOCdateString('ComparisonCompDate'),
                        
                        'I1': self.noOfParcels,                #PARCELS COLUMN
                        'I2': self.ammoItems,
                        'I3': self.ammoItems,
                        'I4': self.totalItemsNos,
                        'I5': "",
                        'I6': "",
                        'A' : self.analyst,                          # NAMES & INITIALS COLUMN
                        'B' : self.analyst,
                        'C': self.Balscanner,
                        'D': self.Balscanner,
                        'E': self.analyst,
                        'F' : self.analyst,
                        'G' : "CPR",
                        'H' : "",
                        'I': "",
                        'J': "",
                        'K': "",
                        'T1': self.CoCDF.getCOCdateString('frmGRLDate'),           # DATE & TIME COLUMN
                        'T2': self.CoCDF.getCOCdateString('BalScanStartDate'),
                        'T3': self.CoCDF.getCOCdateString('BalScanCompDate'),
                        'T4': self.CoCDF.getCOCdateString('toCPRDate'),
                        'T5': "",
                        'T6': "",
                        'P1': "CaseWork",            # PURPOSE COLUMN
                        'P2': "BalScan",
                        'P3': "BalScan Done",
                        'P4': "Case DOne",
                        'P5': "",
                        'P6': "",
                    }
        elif(x==3):                                     # TEAM WITH OUT BALSCAN
            return {   
                        'AGENCY_CASE' : self.fullCaseNumber,
                        'AGENCY_CASE2' : self.AdditionalCaseNumbers,
                        'ANALYST': self.analyst,
                        'REVIEWER': self.reviewer,
                        'PARCEL_DETAILS': EvDetails,
                        'REVIEW_START': self.CoCDF.getCOCdateString('ReviewStartDate'),
                        'REVIEW_END': self.CoCDF.getCOCdateString('ReviewEndDate'),
                        'COMP_END': self.CoCDF.getCOCdateString('ComparisonCompDate'),                        

                        'I1': self.noOfParcels,                #PARCELS COLUMN
                        'I2': self.ammoItems,
                        'I3': self.ammoItems,
                        'I4': self.totalItemsNos,
                        'I5': "",
                        'I6': "",
                        'A' : self.TeamMember,                          # NAMES & INITIALS COLUMN
                        'B' : self.TeamMember,
                        'C': self.analyst,
                        'D': self.analyst,
                        'E': self.TeamMember,
                        'F' : self.TeamMember,
                        'G' : "CPR",
                        'H' : "",
                        'I': "",
                        'J': "",
                        'K': "",
                        'T1': self.CoCDF.getCOCdateString('frmGRLDate'),           # DATE & TIME COLUMN
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
            return  {   
                        'AGENCY_CASE' : self.fullCaseNumber,
                        'AGENCY_CASE2' : self.AdditionalCaseNumbers,
                        'ANALYST': self.analyst,
                        'REVIEWER': self.reviewer,
                        'PARCEL_DETAILS': EvDetails,
                        'REVIEW_START': self.CoCDF.getCOCdateString('ReviewStartDate'),
                        'REVIEW_END': self.CoCDF.getCOCdateString('ReviewEndDate'),
                        'COMP_END': self.CoCDF.getCOCdateString('ComparisonCompDate'),


                        'I1': self.noOfParcels,                #PARCELS COLUMN
                        'I2': self.ammoItems,
                        'I3': self.ammoItems,
                        'I4': self.ammoItems,
                        'I5': self.ammoItems,
                        'I6': self.totalItemsNos,
                        'A' : self.TeamMember,                          # NAMES & INITIALS COLUMN
                        'B' : self.TeamMember,
                        'C': self.analyst,
                        'D': self.analyst,
                        'E': self.TeamMember,
                        'F' : self.TeamMember,
                        'G' : self.Balscanner,
                        'H' : self.Balscanner,
                        'I': self.TeamMember,
                        'J': self.TeamMember,
                        'K': "CPR",
                        'T1': self.CoCDF.getCOCdateString('frmGRLDate'),           # DATE & TIME COLUMN
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

    # Concatenate Items Details in a single String
    def parcelDetailsStringMaker(self, ParcelNo, Quantity, EVCaliber, EVDetails, ItemsNo, notes):
        if(notes == None or notes == "None"):
            notes = ""
        
        # # converts digit to text
        # inflectEngine = inflect.engine()
        # Q = inflectEngine.number_to_words(Quantity)

        if(ParcelNo == "and"):
            return (f"and {self.numberToWord(Quantity)} {str(EVCaliber)} caliber {EVDetails} (Item {ItemsNo}) {notes}")
        else:
            return (f"Parcel {str(ParcelNo)} : {self.numberToWord(Quantity)} {str(EVCaliber)} caliber {EVDetails} (Item {ItemsNo}) {notes}")

    # gets LIST of parcels in case and used it to combine in a single string.
    def getAndSetParcels(self):
        parcels = self.ParcelsDF.getParcelsDetailsForProcessingSheet()

        caseDetailsList = []

        for index, item in enumerate(parcels):
            logging.info(f"index {index}")
            # for parcel add PARCEL 1 to start
            if(index == 0):
                caseDetailsList.append(self.parcelDetailsStringMaker(
                                        ParcelNo=item[0], Quantity=item[5], EVCaliber=item[1],
                                        EVDetails=item[3], ItemsNo=item[4], notes=item[6]))
            else:
                # gets the parcel No of previous parcel
                oldParcelNo = parcels[index-1][0]

                if(item[0] != oldParcelNo):
                    caseDetailsList.append(self.parcelDetailsStringMaker(
                                            ParcelNo=item[0], Quantity=item[5], EVCaliber=item[1], 
                                            EVDetails=item[3], ItemsNo=item[4], notes=item[6]))
                else:
                    caseDetailsList.append(self.parcelDetailsStringMaker(
                                            ParcelNo="and", Quantity=item[5], EVCaliber=item[1], 
                                            EVDetails=item[3], ItemsNo=item[4], notes=item[6]))
        # this method also joins parcel string and returns a single string of case details for 
        # processing sheet
        logging.info(caseDetailsList)
        return ("").join(caseDetailsList)

    # Poplulate and Generate Processing Sheet
    def proceesingSheetMaker(self):
        context = self.setCoCandEVdetails()
        self.processingDocTemplate.render(context)
        self.processingDocTemplate.save(os.path.join(self.currentCaseFolderPath,
                                        f'1. Processing Sheet-{self.ftmNumber}.docx'))


class FirearmsProcessor(Sheets):

    def __init__(self, ftmNumber) -> None:
        super().__init__(ftmNumber)
            
        # List of firearms 
        self.firearms = self.ParcelsDF.getFirearmsOrAmmoDF(firearms).sort_values('ParcelNo').values.tolist()
        # Create instance of DOCX TEMPLATE
        self.firearmsDocTemplate = DocxTemplate(firearmsTemplatePath)

    def testFiresFromItemNo(self, itemNo: str) -> str:
        itemsToCheck = {
                        'R': f'{itemNo}TC1 & {itemNo}TC2',
                        'P': f'{itemNo}TC1 & {itemNo}TC2',
                        'S': f'{itemNo}TS1 & {itemNo}TS2',
                        'M': f'{itemNo}TC1 & {itemNo}TC2',
                        }
        logging.info(f"Item No is {itemNo}")

        if itemNo == None or "":
            return ""
        else:
            itemNo = itemNo.upper()
            # loops through dictionary to find respective item test fires names
            for key in itemsToCheck:
                if(itemNo.find(key) != -1):
                    logging.info(f"Test fires from firearm {itemsToCheck[key]}")
                    return itemsToCheck[key]


    # Iterate through each firearm in firarsm List and save a worksheet with corresponding item No
    def firearmSheetMaker(self):
        if len(self.firearms) > 0: 
            for firearm in self.firearms:

                testFires = self.testFiresFromItemNo(firearm[4])
                yearShort = str(self.caseNumberParts[0])

                context =   {   
                                'AGENCY_CASE' : self.fullCaseNumber,
                                'AGENCY_CASE2' : self.AdditionalCaseNumbers,
                                'ITEM': firearm[4],
                                'EXAMINER': self.analyst,
                                'REVIEWER': self.reviewer,
                                'DATE' : self.processingDate,
                                'CALIBER' : firearm[1],
                                'FTMNO' : self.caseNumberParts[2],
                                'MARKING': str(firearm[4])+"/"+str(self.caseNumberParts[2])+"/"+yearShort[2:],
                                'ABIS': self.BalscanDate,
                                'TESTFIRES': testFires
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
        self.cartridges = self.ParcelsDF.getFirearmsOrAmmoDF(cartridge).sort_values('ParcelNo').values.tolist()
        # Create instance of DOCX TEMPLATE
        self.cartridgeTemplate = DocxTemplate(cartridgeTemplatePath)


    # Iterate through each firearm in firarsm List and save a worksheet with corresponding item No
    def cartridgeSheetMaker(self):
        context =   {   
                        'AGENCY_CASE' : self.fullCaseNumber,
                        'AGENCY_CASE2' : self.AdditionalCaseNumbers,
                        'EXAMINER': self.analyst,
                        'REVIEWER': self.reviewer,
                        'DATE' : self.processingDate,
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
        self.bullets = self.ParcelsDF.getFirearmsOrAmmoDF(bullet).sort_values('ParcelNo').values.tolist()
        # Create instance of DOCX TEMPLATE
        self.bulletDocTemplate = DocxTemplate(bulletTemplatePath)


    # Iterate through each firearm in firarsm List and save a worksheet with corresponding item No
    def bulletSheetMaker(self):
        # for bullet in self.bullets:
        #     # yearShort = str(self.caseNumberParts[0])

            context =   {   
                            'AGENCY_CASE' : self.fullCaseNumber,
                            'AGENCY_CASE2' : self.AdditionalCaseNumbers,
                            'EXAMINER': self.analyst,
                            'REVIEWER': self.reviewer,
                            'DATE' : self.processingDate,
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
        self.district = self.ParcelsDF.getValuefrmParcels('District', indexNumber=0)
        self.testStatement = self.caseDetailsDF.getValuefrmCaseDetails(columnName="TestsRequest")

        self.parcels = self.ParcelsDF.getParcelDetailsForReport()



    def reportGenerator(self):
        

        testReport = Report()
        testReport.PageLayout('A4')

        # testReport.add_styles()
        testReport.paraTOD()

        testReport.tableCaseDetails(caseNo1= self.fullCaseNumber, caseNo2=self.AdditionalCaseNumbers, 
                                        addressee=self.addressee, district=self.district)

        testReport.paraEvDetail(Addressee= self.addressee, items=self.numberOfParcels, testRequest=self.testStatement)
        
        testReport.tableEvDetails(self.parcels)

        testReport.tableAnalysisDetails(startDate=self.processingDate, endDate=self.toCPRdate)

        testReport.paraResults()
        testReport.paraNotes()
        testReport.paraDisposition()
        # testReport.footer()
        testReport.save(os.path.join(self.currentCaseFolderPath, f'Report {self.ftmNumber}.docx'))
        
        os.system(f"start {self.currentCaseFolderPath}")
        





if __name__ == "__main__":

    # r = ReportProcessor(123456)
    # r.reportGenerator()
    
    # cpr = CPRProcessor('28/02/2022')
    # cpr.FileCPRMaker()


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
   
    # p = ProcessingSheetProcessor(123456)
    # print(p.currentCaseFolderPath)
    # print(p.proceesingSheetMaker())
    # p.proceesingSheetMaker(UserPaths.checkNcreateUserCaseWorkFolder())

    f = FirearmsProcessor(123456)
    print(f.currentCaseFolderPath)
    print(f.firearmSheetMaker())

    # c = CartridgeProcessor(123456)
    # print(c.cartridges)
    # c.cartridgeSheetMaker()

    # b = BulletProcessor(123456)
    # b.bulletSheetMaker()

    # print(UserPaths().checkNcreateCaseWorkDirectory())
    # os.system(f"start {r.currentCaseFolderPath}")

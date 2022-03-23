from heapq import merge
import os

import inflect
from docxtpl import DocxTemplate
from CusPath import UserPaths

from AccessToDF import CaseDetailsDF, CoCDF, ParcelsDF
from AccessToDF import IdentifiersDF

from identifierDocx import IdentifiersDocument

processingTemplatePath = os.path.join(os.getcwd(), "modules\\templates\\processing.docx")
firearmsTemplatePath = os.path.join(os.getcwd(), "modules\\templates\\firearms.docx")
cartridgeTemplatePath = os.path.join(os.getcwd(), "modules\\templates\\cartridge.docx")
bulletTemplatePath = os.path.join(os.getcwd(), "modules\\templates\\bullet.docx")

DateFormat = "%d-%m-%Y"

firearms = ['rifle', 'pistol', 'shotgun', 'machine pistol']
cartridge = ['cartridge case', 'cartridge cases','cartridge', 'shotshell case', 'shotshell cases', 'shotshell']
bullet = ['bullet', 'metal piece','bullets', 'metal pieces']


class IdentifiersProcessor():
    def __init__(self, batchDate) -> None:
        self.batchDate = batchDate


        # List of Identifiers from dataframe
        self.Identifiers = IdentifiersDF(self.batchDate).identifiersDF.values.tolist()
        
        

    def FileIdentifierMaker(self, saveLocation):
        i = IdentifiersDocument()
        i.PageLayout('A4')
        i.add_styles()
        i.createTwoColumnsPage()

        for identifier in self.Identifiers:
            print(identifier)

            caseNoFull = "PFSA" + str(identifier[1]) + "-" + str(identifier[2]) + "-FTM-" + str(identifier[3]) 

            # i.tableIdentifiersFiles("PFSA2020-123456-FTM-123456", "PFSA2020-123456-FTM-123456", 1, "123 (XX.XX.XXXX)", "ABC&XYZ")
            i.addFileIdentifiers(caseNo1=caseNoFull, caseNo2=str(identifier[5]), parcels=str(identifier[10]),
                                fir=str(identifier[6]), firDate=identifier[7], ps=str(identifier[8]),
                                district=str(identifier[9]))

        i.saveDoc(os.path.join(saveLocation, f"Identifiers.docx"))

    def EnvelopsMaker(self, saveLocation):
        i = IdentifiersDocument()
        i.PageLayout('A4')
        i.add_styles()
        i.createTwoColumnsPage()

        for envelop in self.Identifiers:
            print(envelop)

            caseNoFull = "PFSA" + str(envelop[1]) + "-" + str(envelop[2]) + "-FTM-" + str(envelop[3]) 

            # i.tableIdentifiersFiles("PFSA2020-123456-FTM-123456", "PFSA2020-123456-FTM-123456", 1, "123 (XX.XX.XXXX)", "ABC&XYZ")
            i.addEnvelopsIdentifiers(caseNo1=caseNoFull, AddressTo=envelop[4],district=str(envelop[9]) )

        i.saveDoc(os.path.join(saveLocation, "Envelops.docx"))


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
        self.AdditionalCaseNumbers = self.caseDetailsDF.getValuefrmCaseDetails(columnName="CaseNosAddl")
        self.analyst = self.caseDetailsDF.getValuefrmCaseDetails(columnName="AnalystName")
        self.reviewer = self.caseDetailsDF.getValuefrmCaseDetails(columnName="ReviewerName")
        self.processingDate = self.CoCDF.getCOCdate('ProcessingDate')
        self.BalscanDate = self.CoCDF.getCOCdateString("BalScanCompDate")


    def fullCaseNumber(self) -> str:
        x = self.caseDetailsDF.getCaseNoParts()
        return "PFSA"+ str(x[0]) + "-" + str(x[1]) + "-FTM-" + str(x[2])

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
                        'P4': "Case DOne",
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
        
        # converts digit to text
        inflectEngine = inflect.engine()
        Q = inflectEngine.number_to_words(Quantity)

        if(ParcelNo == "and"):
            return f"and {Q} {str(EVCaliber)} caliber {EVDetails} (Item {ItemsNo}) {notes}"
        else:
            return f"Parcel {str(ParcelNo)} : {Q} {str(EVCaliber)} caliber {EVDetails} (Item {ItemsNo}) {notes}"



    # gets LIST of parcels in case and used it to combine in a single string.
    def getAndSetParcels(self):
        parcels = self.ParcelsDF.getParcelsDetailsForProcessingSheet()
        print(parcels)
        caseDetailsList = []

        if(len(parcels) == 0):
            print("No parcels Submitted")
            return ""

        else:
            pass


        for index, item in enumerate(parcels):
            # for parcel add PARCEL 1 to start
            if(index == 0):
                caseDetailsList.append(self.parcelDetailsStringMaker(ParcelNo=item[0], Quantity=item[5], EVCaliber=item[1]
                                                , EVDetails=item[3], ItemsNo=item[4], notes=item[6]))
            else:
                # gets the parcel No of previous parcel
                oldParcelNo = parcels[index-1][0]

                if(item[0] != oldParcelNo):
                    caseDetailsList.append(self.parcelDetailsStringMaker(ParcelNo=item[0], Quantity=item[5], EVCaliber=item[1]
                                                    , EVDetails=item[3], ItemsNo=item[4], notes=item[6]))
                else:
                    caseDetailsList.append(self.parcelDetailsStringMaker(ParcelNo="and", Quantity=item[5], EVCaliber=item[1]
                                                    , EVDetails=item[3], ItemsNo=item[4], notes=item[6]))
                
        return (", ").join(caseDetailsList)

    def proceesingSheetMaker(self, saveLocation):
        context = self.setCoCandEVdetails()

        

        # context = contextCOC.update(contextMain)
        # print(contextMain)
        self.processingDocTemplate.render(context)
        self.processingDocTemplate.save(os.path.join(saveLocation, f'1. Processing Sheet-{self.ftmNumber}.docx'))


class FirearmsProcessor(Sheets):

    def __init__(self, ftmNumber) -> None:
        super().__init__(ftmNumber)
            
        # List of firearms 
        self.firearms = self.ParcelsDF.getFirearmsOrAmmoDF(firearms).sort_values('ParcelNo').values.tolist()
        # Create instance of DOCX TEMPLATE
        self.firearmsDocTemplate = DocxTemplate(firearmsTemplatePath)


    # Iterate through each firearm in firarsm List and save a worksheet with corresponding item No
    def firearmSheetMaker(self, saveLocation):
        if len(self.firearms) > 0: 
            for firearm in self.firearms:
                yearShort = str(self.caseNumberParts[0])

                context =   {   
                                'AGENCY_CASE' : self.fullCaseNumber,
                                'AGENCY_CASE2' : self.AdditionalCaseNumbers,
                                'ITEM': firearm[4],
                                'EXAMINER': self.analyst,
                                'REVIEWER': self.reviewer,
                                'DATE' : self.processingDate.strftime(DateFormat),
                                'CALIBER' : firearm[1],
                                'FTMNO' : self.caseNumberParts[2],
                                'MARKING': str(firearm[4])+"/"+str(self.caseNumberParts[1])+"/"+yearShort[2:],
                                'ABIS': self.BalscanDate,


                            }

                self.firearmsDocTemplate.render(context)
                self.firearmsDocTemplate.save(os.path.join(saveLocation,
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
    def cartridgeSheetMaker(self, saveLocation):
        context =   {   
                        'AGENCY_CASE' : self.fullCaseNumber,
                        'AGENCY_CASE2' : self.AdditionalCaseNumbers,
                        'EXAMINER': self.analyst,
                        'REVIEWER': self.reviewer,
                        'DATE' : self.processingDate.strftime(DateFormat),
                        # 'CALIBER' : cartridge[1],
                        # 'FTMNO' : self.caseNumberParts[2],
                        # 'MARKING': str(cartridge[4])+"/"+str(self.caseNumberParts[1])+"/"+yearShort[2:],
                        # 'ABIS': self.BalscanDate.strftime(DateFormat),
                    }
        self.cartridgeTemplate.render(context)
        self.cartridgeTemplate.save(os.path.join(saveLocation,
                                        f"3. cartridge{self.ftmNumber}.docx"))


class BulletProcessor(Sheets):

    def __init__(self, ftmNumber) -> None:
        super().__init__(ftmNumber)
            
        # List of firearms 
        self.bullets = self.ParcelsDF.getFirearmsOrAmmoDF(bullet).sort_values('ParcelNo').values.tolist()
        # Create instance of DOCX TEMPLATE
        self.bulletDocTemplate = DocxTemplate(bulletTemplatePath)


    # Iterate through each firearm in firarsm List and save a worksheet with corresponding item No
    def bulletSheetMaker(self, saveLocation):
        # for bullet in self.bullets:
        #     # yearShort = str(self.caseNumberParts[0])

            context =   {   
                            'AGENCY_CASE' : self.fullCaseNumber,
                            'AGENCY_CASE2' : self.AdditionalCaseNumbers,
                            'EXAMINER': self.analyst,
                            'REVIEWER': self.reviewer,
                            'DATE' : self.processingDate.strftime(DateFormat),
                            # 'CALIBER' : cartridge[1],
                            # 'FTMNO' : self.caseNumberParts[2],
                            # 'MARKING': str(cartridge[4])+"/"+str(self.caseNumberParts[1])+"/"+yearShort[2:],
                            # 'ABIS': self.BalscanDate.strftime(DateFormat),
                        }
            self.bulletDocTemplate.render(context)
            self.bulletDocTemplate.save(os.path.join(saveLocation,
                                            f"4. bullet-{self.ftmNumber}.docx"))
        
        




if __name__ == "__main__":

    i = IdentifiersProcessor("1/3/2022")
    i.FileIdentifierMaker(UserPaths().checkNcreateCaseWorkDirectory())
    i.EnvelopsMaker(UserPaths().checkNcreateCaseWorkDirectory())
   
    p = ProcessingSheetProcessor(123456)
    print(p.proceesingSheetMaker(UserPaths().checkNcreateCaseWorkDirectory()))
    # p.proceesingSheetMaker(UserPaths.checkNcreateUserCaseWorkFolder())

    f = FirearmsProcessor(123456)
    print(f.firearms)
    print(f.firearmSheetMaker(UserPaths().checkNcreateCaseWorkDirectory()))

    c = CartridgeProcessor(123456)
    # print(c.cartridges)
    c.cartridgeSheetMaker(UserPaths().checkNcreateCaseWorkDirectory())

    b = BulletProcessor(123456)
    b.bulletSheetMaker(UserPaths().checkNcreateCaseWorkDirectory())

    print(UserPaths().checkNcreateCaseWorkDirectory())

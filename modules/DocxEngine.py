import os
from datetime import datetime

from docxtpl import DocxTemplate
import pandas as pd
from CusPath import UserPaths

from AccessToDF import CaseDetailsDF as CaseDetails, CoCDF as COC, ParcelsDF as Parcels
from AccessToDF import IdentifiersDF

from identifierDocx import IdentifiersDocument

firearmsTemplatePath = os.path.join(os.getcwd(), "modules\\templates\\firearms.docx")
cartridgeTemplatePath = os.path.join(os.getcwd(), "modules\\templates\\firearms.docx")
bulletTemplatePath = os.path.join(os.getcwd(), "modules\\templates\\bullet.docx")

DateFormat = "%d-%m-%Y"

firearms = ['rifle', 'pistol', 'shotgun', 'machine pistol']
cartridge = ['cartridge case']
bullet = ['bullet', 'metal piece']

class Sheets():

    # Gets all required data
    def __init__(self, ftmNumber) -> None:
        self.ftmNumber = ftmNumber

        # create instance of Tables to be used in all sheets processor
        self.caseDetails = CaseDetails(self.ftmNumber)
        self.COC = COC(self.ftmNumber)
        self.Parcels = Parcels(self.ftmNumber)

        # These variables will be used in all worksheets
        self.fullCaseNumber = self.fullCaseNumber()
        self.caseNumberParts = self.caseDetails.getCaseNoParts()
        self.AdditionalCaseNumbers = self.caseDetails.getValuefrmCaseDetails(columnName="CaseNosAddl")
        self.analyst = self.caseDetails.getValuefrmCaseDetails(columnName="AnalystName")
        self.reviewer = self.caseDetails.getValuefrmCaseDetails(columnName="ReviewerName")

        self.processingDate = self.COC.getCOCdate('ProcessingDate')
        self.BalscanDate = self.COC.getCOCdate("BalScanCompDate")


    def fullCaseNumber(self) -> str:
        x = self.caseDetails.getCaseNoParts()
        return "PFSA"+ str(x[0]) + "-" + str(x[1]) + "-FTM-" + str(x[2])
    

class FirearmsProcessor(Sheets):

    def __init__(self, ftmNumber) -> None:
        super().__init__(ftmNumber)
            
        # List of firearms 
        self.firearms = self.Parcels.getFirearmsOrAmmoDF(firearms).sort_values('ParcelNo').values.tolist()
        # Create instance of DOCX TEMPLATE
        self.firearmsDocTemplate = DocxTemplate(firearmsTemplatePath)


    # Iterate through each firearm in firarsm List and save a worksheet with corresponding item No
    def firearmSheetMaker(self):
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
                                'ABIS': self.BalscanDate.strftime(DateFormat),

                            }
                self.firearmsDocTemplate.render(context)
                self.firearmsDocTemplate.save(os.path.join(UserPaths().checkNcreateUserCaseWorkFolder(), f"{self.caseNumberParts[2]}-{firearm[0]}-firearms.docx"))
        else:
            print("No firearms sheet is generated as no data is passed to processor")


class CartridgeProcessor(Sheets):

    def __init__(self, ftmNumber) -> None:
        super().__init__(ftmNumber)
            
        # List of firearms 
        self.cartridges = self.Parcels.getFirearmsOrAmmoDF(cartridge).sort_values('ParcelNo').values.tolist()
        # Create instance of DOCX TEMPLATE
        self.firearmsDocTemplate = DocxTemplate(cartridgeTemplatePath)


    # Iterate through each firearm in firarsm List and save a worksheet with corresponding item No
    def cartridgeSheetMaker(self):
        if len(self.cartridges) > 0:
            for cartridge in self.cartridges:
                # yearShort = str(self.caseNumberParts[0])

                context =   {   
                                'AGENCY_CASE' : self.fullCaseNumber,
                                'AGENCY_CASE2' : self.AdditionalCaseNumbers,
                                'ITEM': cartridge[4],
                                'EXAMINER': self.analyst,
                                'REVIEWER': self.reviewer,
                                'DATE' : self.processingDate.strftime(DateFormat),
                                # 'CALIBER' : cartridge[1],
                                # 'FTMNO' : self.caseNumberParts[2],
                                # 'MARKING': str(cartridge[4])+"/"+str(self.caseNumberParts[1])+"/"+yearShort[2:],
                                # 'ABIS': self.BalscanDate.strftime(DateFormat),
                            }
                self.firearmsDocTemplate.render(context)
                self.firearmsDocTemplate.save(os.path.join(UserPaths.userCaseWorkFolder(),
                                                f"{self.caseNumberParts[2]}-{cartridge[0]}-cartridge.docx"))
        else:
            print("No cartridge sheet is generated as no data is passed to processor")


class BulletProcessor(Sheets):

    def __init__(self, ftmNumber) -> None:
        super().__init__(ftmNumber)
            
        # List of firearms 
        self.bullets = self.Parcels.getFirearmsOrAmmoDF(bullet).sort_values('ParcelNo').values.tolist()
        # Create instance of DOCX TEMPLATE
        self.firearmsDocTemplate = DocxTemplate(cartridgeTemplatePath)


    # Iterate through each firearm in firarsm List and save a worksheet with corresponding item No
    def cartridgeSheetMaker(self):
        for bullet in self.bullets:
            # yearShort = str(self.caseNumberParts[0])

            context =   {   
                            'AGENCY_CASE' : self.fullCaseNumber,
                            'AGENCY_CASE2' : self.AdditionalCaseNumbers,
                            'ITEM': bullet[4],
                            'EXAMINER': self.analyst,
                            'REVIEWER': self.reviewer,
                            'DATE' : self.processingDate.strftime(DateFormat),
                            # 'CALIBER' : cartridge[1],
                            # 'FTMNO' : self.caseNumberParts[2],
                            # 'MARKING': str(cartridge[4])+"/"+str(self.caseNumberParts[1])+"/"+yearShort[2:],
                            # 'ABIS': self.BalscanDate.strftime(DateFormat),
                        }
            self.firearmsDocTemplate.render(context)
            self.firearmsDocTemplate.save(os.path.join(UserPaths.userCaseWorkFolder(),
                                            f"{self.caseNumberParts[2]}-{bullet[0]}-bullet.docx"))
        
        
class IdentifiersProcessor():
    def __init__(self, batchDate) -> None:
        self.batchDate = batchDate

        # Instanace of Identifiers Dataframe
        self.IdentifiersDF = IdentifiersDF(self.batchDate)

        # List of Identifiers from dataframe
        self.identifiers = self.IdentifiersDF.combineCaseDetailsWithFIRDate()
        
        



    def FileIdentifierMaker(self):
        i = IdentifiersDocument()
        i.PageLayout('A4')
        i.add_styles()
        i.createTwoColumnsPage()

        for identifier in self.identifiers:
            print(identifier)

            caseNoFull = "PFSA" + str(identifier[1]) + "-" + str(identifier[2]) + "-FTM-" + str(identifier[3]) 

            # i.tableIdentifiersFiles("PFSA2020-123456-FTM-123456", "PFSA2020-123456-FTM-123456", 1, "123 (XX.XX.XXXX)", "ABC&XYZ")
            i.addFileIdentifiers(caseNo1=caseNoFull, caseNo2=str(identifier[5]), parcels=str(identifier[10]),
                                fir=str(identifier[6]), firDate=identifier[11], ps=str(identifier[8]),
                                district=str(identifier[9]))

        i.saveDoc(UserPaths().CurrentCaseWorkFolder)




if __name__ == "__main__":
    f = FirearmsProcessor(123456)
    print(f.firearms)
    print(f.firearmSheetMaker())

    i = IdentifiersProcessor("1/3/2022")
    i.FileIdentifierMaker()
   

import os
import datetime
import configparser
from pathlib import Path
from math import ceil


class UserPaths():

    userHomePath = os.path.expanduser("~")
    userDesktopCaseworkPath = os.path.join(userHomePath,"Desktop", "Casework")
    processingTemplatePath = os.path.join(os.getcwd(), "templates\\processing.docx")
    firearmsTemplatePath = os.path.join(os.getcwd(), "templates\\firearms.docx")
    cartridgeTemplatePath = os.path.join(os.getcwd(), "templates\\cartridge.docx")
    bulletTemplatePath = os.path.join(os.getcwd(), "templates\\bullet.docx")
    reportTemplatePath = './templates/template.docx'
    cprTemplatePath = './templates/cpr.docx'


    def __init__(self) -> None:
        self.CurrentCaseWorkParentFolder = self.checkNcreateCaseWorkDirectory()

    @classmethod
    def checkNcreateFolder(cls, path):
        if os.path.isdir(path):
            return path
        else:
            Path(path).mkdir(parents=True, exist_ok=True)
            return path

    # CHeck if a CASEWORK folder exist. if None then create a casework directory on desktop
    # and then return the path.

    @classmethod
    def checkNcreateCaseWorkDirectory(cls):
        # Load configuration from configuration.ini
        config = configparser.ConfigParser()
        config.read("configuration.ini")
        # Check if directory exists in the custom path specified in the configuration file
        custom_path = config.get("Paths", "CaseworkPath")
        if os.path.isdir(custom_path):
            return custom_path
        # Define the list of paths to check
        paths_to_check = [
            "E:\\Casework",
            "D:\\Casework",
            cls.userDesktopCaseworkPath
        ]
        # Check each path and return the first one that exists
        for path in paths_to_check:
            if os.path.isdir(path):
                return path

        # If none of the directories exist, create the default one on the Desktop
        default_path = cls.userDesktopCaseworkPath
        try:
            os.makedirs(default_path, exist_ok=True)
        except OSError as e:
            print(f"Error creating directory {default_path}: {e}")
            raise
        
        return default_path

    def fileWriteableStateCheck(self, filePath):
        if(os.path.isfile(filePath)):
            tempFile = filePath + ".temp"
            try:
                os.rename(filePath, tempFile)
                os.rename(tempFile, filePath)
                print("file is writeable and closed")
                return 'close'
            except OSError:
                print(f'{filePath} is still open. Close this file.')
            finally:
                return 'open'
        else:
            print("File does not exist.")
            return "close"

    @classmethod
    def weekOfCurrentMonth(self):

        dt = datetime.datetime.now()
        first_day = dt.replace(day=1)

        dom = dt.day
        adjusted_dom = dom + first_day.weekday()

        return "Week" + str(int(ceil(adjusted_dom/7.0)))

    def makeFolderfrmDate(self, date: datetime):
        ''' This method creates folder from batch date in casework directory
            and returns path in string format of current batch date.'''
        currentYear = date.strftime("%Y")
        currentMonth = date.strftime("%B")
        batchDate = date.strftime("%d-%m-%Y")

        currentBatchFolder = os.path.join(
            self.CurrentCaseWorkParentFolder, currentYear, currentMonth, batchDate)

        return self.checkNcreateFolder(currentBatchFolder)

    @classmethod
    def makeFolderInPath(cls, path: str, caseNo: str):
        for folder in ['Comparison Pictures', 'Evidence Pictures','Sheets']:
            cls.checkNcreateFolder(os.path.join(path, caseNo, folder))
        return os.path.join(path, caseNo, 'Sheets')


if __name__ == "__main__":
    path = UserPaths()
    print(path.CurrentCaseWorkParentFolder)
    path.fileWriteableStateCheck(
        "C:\\Users\\Faisal\\Desktop\\Casework\\123456-1-firearms.docx")

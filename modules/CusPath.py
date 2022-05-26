import os
import datetime
from math import ceil


class UserPaths():

    userHomePath = os.path.expanduser("~")

    def __init__(self) -> None:
        self.CurrentCaseWorkParentFolder = self.checkNcreateCaseWorkDirectory()

    @classmethod
    def checkNcreateFolder(cls, path):
        if os.path.isdir(path):
            return path
        else:
            os.makedirs(path)
            return path

    # CHeck if a CASEWORK folder exist. if None then create a casework directory on desktop
    # and then return the path.

    @classmethod
    def checkNcreateCaseWorkDirectory(cls):
        if os.path.isdir("E:\Casework") == True:
            return "E:\Casework"
        elif os.path.isdir("D:\Casework") == True:
            return "D:\Casework"
        elif os.path.isdir(os.path.join(cls.userHomePath, "Desktop", "Casework")) == True:
            return os.path.join(cls.userHomePath, "Desktop", "Casework")
        else:
            os.makedirs(os.path.join(cls.userHomePath, "Desktop", "Casework"))
            return os.path.join(cls.userHomePath, "Desktop", "Casework")

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
        caseFolder = os.path.join(path, caseNo, 'Sheets')
        return cls.checkNcreateFolder(caseFolder)


if __name__ == "__main__":
    path = UserPaths()
    print(path.CurrentCaseWorkParentFolder)
    path.fileWriteableStateCheck(
        "C:\\Users\\Faisal\\Desktop\\Casework\\123456-1-firearms.docx")

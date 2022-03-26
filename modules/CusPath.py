import os
import datetime
from math import ceil

class UserPaths():

    userHomePath = os.path.expanduser("~")

    def __init__(self) -> None:
        self.CurrentCaseWorkFolder = self.checkNcreateCaseWorkDirectory()
        self.CurrentWeekFolder = self.makeCurrentWeekFolder()

    # CHeck if a CASEWORK folder exist. if None then create a casework directory on desktop
    # and then return the path.
    def checkNcreateCaseWorkDirectory(self):
        if os.path.isdir("E:\Casework") == True:
            return "E:\Casework"
        elif os.path.isdir("D:\Casework") == True:
            return "D:\Casework"
        elif os.path.isdir(os.path.join(self.userHomePath,"Desktop", "Casework")) == True:
            return os.path.join(self.userHomePath,"Desktop", "Casework")
        else:
            os.makedirs(os.path.join(self.userHomePath,"Desktop", "Casework"))
            return os.path.join(self.userHomePath,"Desktop", "Casework")

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

    def week_of_month(self):

        dt = datetime.datetime.now()
        first_day = dt.replace(day=1)

        dom = dt.day
        adjusted_dom = dom + first_day.weekday()

        return "Week" + str(int(ceil(adjusted_dom/7.0)))

    def makeCurrentWeekFolder(self):
        currentYear = datetime.datetime.now().strftime("%Y")
        currentMonth = datetime.datetime.now().strftime("%B")
        currentWeek = self.week_of_month()

        currentWeekFolder = os.path.join(self.CurrentCaseWorkFolder, currentYear, currentMonth, currentWeek)

        if os.path.isdir(currentWeekFolder):
            return currentWeekFolder
        else:
            os.makedirs(currentWeekFolder)
            return currentWeekFolder

if __name__ == "__main__":
    path = UserPaths()
    print(path.CurrentCaseWorkFolder)
    path.fileWriteableStateCheck("C:\\Users\\Faisal\\Desktop\\Casework\\123456-1-firearms.docx")
    print('Test')
    print(path.CurrentWeekFolder)

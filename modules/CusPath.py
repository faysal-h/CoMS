import os

class UserPaths():

    userHomePath = os.path.expanduser("~")

    def __init__(self) -> None:
        self.CurrentCaseWorkFolder = self.checkNcreateCaseWorkDirectory()

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
            except OSError:
                print(f'{filePath} is still open. Close this file.')
        else:
            print("File does not exist.")

    @classmethod
    def checkNcreateUserCaseWorkFolder(cls) -> str :
        if os.path.isdir("E:\Casework") == True:
            return "E:\Casework"
        else:
            return os.path.join(cls.userHomePath, "Desktop", "CaseWork")




if __name__ == "__main__":
    path = UserPaths()
    print(path.CurrentCaseWorkFolder)
    path.fileWriteableStateCheck("C:\\Users\\Faisal\\Desktop\\Casework\\123456-1-firearms.docx")

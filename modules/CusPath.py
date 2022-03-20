import os

from access import CaseDetails

class UserPaths():

    userHomePath = os.path.expanduser("~")

    def __init__(self) -> None:
        self.caseWorkDirectory()

    # CHeck if a CASEWORK folder exist. if None then create a casework directory on desktop
    def caseWorkDirectory(self):
        if os.path.isdir("E:\Casework") == True:
            pass
        elif os.path.isdir("E:\Casework") == True:
            pass
        elif os.path.isdir(os.path.join(self.userHomePath,"Desktop", "Casework")) == True:
            pass
        else:
            os.makedirs(os.path.join(self.userHomePath,"Desktop", "Casework"))


    @classmethod
    def userCaseWorkFolder(cls) -> str :
        if os.path.isdir("E:\Casework") == True:
            return "E:\Casework"
        else:
            return os.path.join(cls.userHomePath, "Desktop", "CaseWork")




if __name__ == "__main__":
    path = UserPaths()
    print(UserPaths.userCaseWorkFolder())
    print(UserPaths.userHomePath)
    
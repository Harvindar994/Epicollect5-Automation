import pyautogui as pt
from time import sleep
import clipboard
import re
# while True:
#     posXY = pt.position()
#     print(posXY, pt.pixel(posXY[0], posXY[1]))
#     sleep(1)
#     if posXY[0] == 0:
#         break


class AutoRunner:
    def __init__(self):
        self.imgLocater = 'assets/locator.png'
        self.imgPreviewsActive = 'assets/previews_active.png'
        self.imgNextActive = "assets/next_active.png"
        self.imgNext = "assets/next.png"
        self.imgPreviews = "assets/previews.png"
        self.imgTick = 'assets/tick.png'
        self.addBranch = 'assets/add.png'
        self.buttonSaveEntry = 'assets/saveentry.png'

        # add branch buttons
        self.buttonSaveBranch = 'assets/savebranch.png'
        self.buttonSarpanch = 'assets/sarpanch.png'
        self.buttonSarpanchSelected = "assets/sarpanchselected.png"
        self.fieldname = 'assets/namefield.png'
        self.fieldMobile = 'assets/mobilefield.png'
        self.fieldyes = 'assets/yes.png'
        self.field = 'assets/field.png'
        self.add = 'assets/add.png'

        self.villageSelecter = 'assets/villageselect.png'

        # Locaters
        self.addBranchLocater = 'assets/addBranchLocater.png'
        self.keyboardLocater = 'assets/keyboard.png'

        # common buttons
        self.buttonCopy = 'assets/copy.png'

        # Pages
        self.pageDistrict = 'district'
        self.pageBlock = 'block'
        self.pageCluster = 'cluster'
        self.pageFramer = 'farmer'
        self.pageSave = 'save'

        # Branch pages
        self.branchRole = 'role'
        self.branchName = 'name'
        self.branchMobile = 'mobile'
        self.branchSelectSmartphone = 'smartphone'
        self.branchSave = 'save'

        self.pages = {self.pageDistrict:1, self.pageBlock:2, self.pageCluster:3,
                      self.pageFramer:4, self.pageSave: 5}
        self.branchPages = {self.branchRole:6, self.branchName:7, self.branchMobile:8,
                            self.branchSelectSmartphone:9, self.branchSave: 10}

        self.addingBranch = False
        self.currentBranchPage = 1
        self.currentPage = 4

        self.ButtonQuit = 'assets/quit.png'
        self.ButtonQuit2 = 'assets/quit2.png'

    def quitfromBranchFrom(self):
        run = True
        count = 1
        while run:
            if count > 7:
                return False
            position = pt.locateCenterOnScreen(self.buttonSarpanch, confidence=.9)
            if position is None:
                position = pt.locateCenterOnScreen(self.buttonSarpanchSelected, confidence=.9)
            if position is None:
                self.clickonPreviews()
                count += 1
            else:
                if self.clickonButton(self.ButtonQuit, retry=3):
                    if self.clickonButton(self.ButtonQuit2, retry=3):
                        self.resetPages()
                        sleep(1)
                        return True
                    else:
                        return False
                else:
                    return False
        return False

    def saveEntry(self):
        if self.gotoPage(self.pageSave):
            if self.clickonButton(self.buttonSaveEntry, retry=3):
                return True
        return False

    def resetPages(self):
        self.currentPage = self.pages[self.pageFramer]
        self.currentBranchPage = self.branchPages[self.branchRole]
        self.addingBranch = False

    def isKeyboardActive(self):
        position = pt.locateCenterOnScreen(self.keyboardLocater, confidence=.9)
        if position is not None:
            return True
        return False

    def fillBranchFrom(self, name='sdsdf', mobile="3434343434"):
        self.movetoLocater()
        if not self.clickonButton(auto.buttonSarpanch):
            print("Unable to check the Sarpanch radio Button")
            return False
        if not self.clickonNext():
            print("Unable to click on next button")
            return False
        if not self.clickonButton(self.fieldname):
            if not self.clickonButton(self.field):
                print("Unable to find the Name Field")
                return False
        clipboard.copy(name)
        pt.hotkey('ctrl','v')
        if not self.clickonNext():
            print("Unable to click on next button")
            return False
        if not self.clickonButton(self.fieldMobile):
            if not self.clickonButton(self.field):
                print("Unable to find the Mobile Field")
                return False
        clipboard.copy(mobile)
        pt.hotkey('ctrl','v')
        if not self.clickonNext():
            print("Unable to click on next button")
            return False
        if not self.clickonButton(self.fieldyes, confidence=0.9, grayscale=False, retry=2):
            print("Unable to Click on Yes Field")
            return False
        if not self.clickonNext():
            print("Unable to click on next button")
            return False
        if not self.clickonButton(self.buttonSaveBranch, confidence=0.9, retry=2):
            print("Unable to click on Save Brach")
            return False
    def selectAll(self):
        pt.press('end')
        pt.hotkey('shiftleft', 'home')

    def chnageVillage(self, newVillage):
        # print(self.addingBranch, self.currentPage, "Info.......")
        # if not self.clickonPreviews():
        #     print("Unable to click on Previews")
        #     return
        if not self.gotoPage(self.pageCluster):
            return False
        # if not self.addingBranch:
        #     if self.currentPage != self.pages[self.pageCluster]:
        #         self.gotoPage(self.pageCluster)
        # else:
        #     return False
        self.moveFormCurrentPosition(30, 30)
        count = 0
        while True:
            if count >= 3:
                if not self.locateToAddBrach():
                    continue
                if not self.gotoPage(self.pageCluster):
                    continue
                self.moveFormCurrentPosition(30, 30)
                count = 0
            position = pt.locateCenterOnScreen(self.villageSelecter, confidence=.9)
            if position is not None:
                pt.moveTo(position[0], position[1]+100, duration=.05)
                while pt.locateCenterOnScreen(self.villageSelecter, confidence=.9) != None:
                    pt.click(button='left')
                break
            else:
                for i in range(15):
                    pt.scroll(-500)
                sleep(1)
                count += 1
        for i in range(3):
            pt.hotkey('ctrl', 'a')
            clipboard.copy(newVillage)
            pt.hotkey('ctrl', 'v')
        sleep(1)
        if not self.clickonNext():
            print("Unable to click on next")
            return False
        return True

    def locateToAddBrach(self):
        run = True
        count = 1
        position = None
        while run:
            if count > 7:
                return False
            position = pt.locateCenterOnScreen(self.buttonSaveEntry, confidence=.9)
            if position is not None:
                if not self.clickonPreviews():
                    return False
                run = False
            else:
                self.clickonNext()
                count += 1

        count = 1
        while position is None and count <= 3:
            position = pt.locateCenterOnScreen(self.addBranch, confidence=.9)
            count += 1
        if position is not None:
            self.resetPages()
            return True
        return False

    def removeCopyifThere(self):
        if not self.clickonButton(self.buttonCopy, retry=1, grayscale=True, confidence=0.8):
            return True
        return False

    def moveFormCurrentPosition(self, pixel_x=0, pixel_y=0):
        x, y = pt.position()
        x += pixel_x
        y += pixel_y
        pt.moveTo(x, y)

    def gotoPage(self, page):
        if self.addingBranch:
            if page in self.branchPages:
                pageNumber = self.branchPages[page]
                while pageNumber != self.currentBranchPage:
                    if self.currentBranchPage < pageNumber:
                        if not self.clickonNext():
                            return False
                    if self.currentBranchPage > pageNumber:
                        if not self.clickonPreviews():
                            return False
                return True
            else:
                return False
        else:
            if page in self.pages:
                pageNumber = self.pages[page]
                while pageNumber != self.currentPage:
                    if self.currentPage < pageNumber:
                        if not self.clickonNext():
                            return False
                    elif self.currentPage > pageNumber:
                        if not self.clickonPreviews():
                            return False
                return True
            else:
                return False

    def clickonButton(self, button, confidence=.9,  grayscale=False, retry=1):
        count = 1
        while count <= retry:
            position = pt.locateCenterOnScreen(button, confidence=confidence, grayscale=grayscale)
            if position is not None:
                pt.moveTo(position[0], position[1], duration=.05)
                pt.click(button='left')
                if button == self.buttonSaveBranch:
                    self.addingBranch = False
                    self.currentBranchPage = self.branchPages[self.branchRole]
                elif button == self.addBranch:
                    self.addingBranch = True
                    self.currentBranchPage = self.branchPages[self.branchRole]
                return True
            count += 1
        return False

    def clickOnAddBranch(self):
        if self.clickonButton(self.addBranch, confidence=.9, retry=2):
            return True
        return False

    def movetoLocater(self):
        position = pt.locateCenterOnScreen('assets/locator.png', confidence=.9)
        if position is not None:
            pt.moveTo(position[0], position[1], duration=.05)
            return True
        return False

    def clickonNext(self, retry=1):
        position = pt.locateCenterOnScreen(self.imgNextActive, confidence=.9)
        if position is not None:
            x, y = position
            pt.moveTo(x, y, duration=.05)
            pt.click(button='left')
            if self.addingBranch:
                self.currentBranchPage += 1
            else:
                self.currentPage += 1
            return True
        else:
            return False

    def clickonPreviews(self, retry=1):
        count = 1
        while count <= retry:
            position = pt.locateCenterOnScreen(self.imgPreviewsActive, confidence=.9)
            if position is not None:
                x, y = position
                pt.moveTo(x, y, duration=.05)
                pt.click(button='left')
                if self.addingBranch:
                    self.currentBranchPage -= 1
                else:
                    self.currentPage -= 1
                return True
            count += 1
        return False

    def isDeactive_Next(self):
        position = pt.locateCenterOnScreen(self.imgNext, confidence=.9)
        if position is not None:
            return True
        return False

    def isDeactive_Previews(self):
        position = pt.locateCenterOnScreen(self.imgPreviews, confidence=.9)
        if position is not None:
            return True
        return False

villages= ['KJHKHKJKJH jHJHGJHGJHGJ', '123123', '3dfdf', 'jYIUYIU', '87032', 'KJHKHKJKJH jHJHGJHGJHGJ']
auto = AutoRunner()
# print(auto.clickonNext())
sleep(2)
pt.hotkey('ctrl','o')
pt.press('win')
# for i in range(3):
#     pt.press('backspace')
#     # pt.hotkey('shift', 'home')

class DataFinder:
    def __init__(self, patterns):
        self.patterns = patterns
        self.compiledPatterns = []
        for pattern in self.patterns:
            compiled = re.compile(pattern)
            self.compiledPatterns.append(compiled)

    def check(self, StringData):
        StringData = str(StringData)
        result = []
        for cp in self.compiledPatterns:
            matches = cp.finditer(StringData)
            for match in matches:
                result.append(match)
        return result


def removeDots(string):
    finalString = string
    for datafilter in [DataFinder([r'[.]*']), DataFinder([r'[`]*']), DataFinder([r'[-]*'])]:
        finder = datafilter
        newstr = ""
        pretart = 0
        matchs = finder.check(finalString)
        if len(matchs) == 0:
            continue
        for match in matchs:
            start, end = match.span()
            newstr += finalString[pretart: start]
            pretart = end
        finalString = newstr

    return finalString


# print(removeDots(".`-Mr. Harvindar Singh."))
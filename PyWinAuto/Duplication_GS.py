import json
import tkinter.messagebox
from collections import namedtuple
from os import remove
from time import sleep
from tkinter import Button, Entry, Frame, Label, StringVar, Tk, X, ttk, Radiobutton

import sys
import clipboard
import mysql.connector
import numpy
import pandas as pd
from mysql.connector import Error
from pandas import ExcelFile
from pywinauto import Application, findwindows, win32defines
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.keyboard import send_keys

from Duplication_GS.classes.MyException import MyException
from Duplication_GS.classes.ProgressBar import ProgressBar

class Duplication_GS(Frame):
    hadDuplicate = False
    nbRowParsitua = 0
    codeSituation = ""
    listCode = []
    IPHost = "192.168.0.3"

    query_count_row = "SELECT COUNT(RECNO) FROM parsitua WHERE RECNO BETWEEN (SELECT RECNO FROM parsitua WHERE CODESITUA=%s) AND (SELECT RECNO FROM parsitua WHERE CODESITUA=%s)"

    def __init__(self, Root):
        Frame.__init__(self, master=Root)

        self.master.title('Duplications : Gestion des situations')
        self.master.resizable(False, False)


        width        = 323
        height       = 95
        # posX         = (widthScreen/2)-(width/2) # NOTE: Horizontal center
        # posY         = (heightScreen/2)-(height/2) # NOTE: Vertical center

        btnInputBoxCodeSituation  = Button(self.master, text="Duplication bloc GS (complète)", command=self.duplicationSituationsAction)
        btnInputBoxCodeCustomGPGT = Button(self.master, text="Duplication bloc GS (partiel)" , command=self.duplicationSituationsMultipleAction)
        btnAffectations           = Button(self.master, text="Affectations"                  , command=self.chGroupeAction)

        # radiovar                        = StringVar(self.master)
        # RadioTest                       = Radiobutton(self.master, text="Radio test", variable=radiovar, value="spam")

        btnInputBoxCodeSituation  .pack(fill=X)
        btnInputBoxCodeCustomGPGT .pack(fill=X)
        btnAffectations           .pack(fill=X)
        # RadioTest                 .pack(      )

        self.master.geometry('%dx%d+%d+%d' % (
            width, height,
            self.master.winfo_screenwidth() - width - 23,
            self.master.winfo_screenheight() - height - 233
        ))
        self.master.wm_attributes("-topmost", 1)
        self.master.protocol(u'WM_DELETE_WINDOW', self.close)

        try: # Watch if already had duplications
            params              = json2obj(open("../var/cache/paramsDuplication.json", "r").read())
            self.hadDuplicate   = True
            self.codeSituation  = params.codeSituation
            self.nbRowParsitua  = params.nbRowParsitua
            self.listCode = params.listCode
        except Exception as e:
            pass

    def close(self):
        try:
            self.RootInputBoxCodeSituation.destroy()
            # self.RootInputBoxCodeSituation.quit()
        except Exception:
            pass

        try:
            self.RootInputBoxCodeCustomGPGT.destroy()
            # self.RootInputBoxCodeCustomGPGT.quit()
        except Exception:
            pass

        try:
            self.master.destroy()
            self.master.quit()
        except Exception:
            pass

# ####################### ACTIONS
    def duplicationSituationsAction(self, listCodeCustomGPGT=[]):
        duplicationUnique = False
        _currentIndex     = 0

        try: # NOTE: Main try/catch
            try: self.windowParamGestionSituation = Application().connect(title_re="Paramètres de la gestion des situations")
            except ElementNotFoundError as exc:
                raise MyException("Fenêtre inexistante", "La fenêtre des gestions des situations n'est pas ouverte")

            if (self.hadDuplicate): raise MyException("Données detectées", "ATTENTION : Il y a eu des duplications, il faut faire les affectations")

            if (not self.inputBoxCodeSituation()): return

            self.listCodeGPGT()

            if (len(listCodeCustomGPGT) > 1): self.sortedListCodeCustom(listCodeCustomGPGT)
            if (not self.getNbLigne()      ): raise MyException("Pas de données trouvées dans MySQL", "Il n'y a pas de situation à dupliquer")

            self.checkCodeSituaStartMatrix()

            if ("UN-" in self.listCode): duplicationUnique = True

            myProgressBar = ProgressBar(self, "Duplications", range(
                len(self.listCode) * self.nbRowParsitua - (0 if duplicationUnique else self.nbRowParsitua)
            ), "Duplication du bloc [{0}-DBM] en cours...".format(self.codeSituation))
            myProgressBar.update()
            # self.master.wait_window(myProgressBar)

            # NOTE: Début Duplication
            for indexGroup in range(len(self.listCode) - (0 if duplicationUnique else 1)):
                for indexRow in range(self.nbRowParsitua):
                    _currentIndex = _currentIndex + 1
                    self.windowParamGestionSituation["Paramètres de la gestion des situations"].Dupliquer.type_keys("{ENTER}")

                    windowConfirm = self._waitExistWindow(winTitle="Confirmation")

                    windowConfirm["Confirmation"].Oui.type_keys("{ENTER}")
                    self.windowParamGestionSituation["Paramètres de la gestion des situations"]["TdxDBGrid1"].set_focus()
                    myProgressBar.num.set(_currentIndex)
                    myProgressBar.subTextVar.set("{0}/{1}".format(_currentIndex, myProgressBar.maximum))
                    myProgressBar.update()

                    if (indexRow < self.nbRowParsitua - 1): self._sendKey("{DOWN}", 0.323, 0.323)

                if (indexGroup < (len(self.listCode) - 1)):
                    for i in range(self.nbRowParsitua - 1): self._sendKey("{UP}", 0.233, 0.233)
            # NOTE: Fin Duplication

            tkinter.messagebox.showinfo("Duplication bloc GS", "Les duplications du bloc [{0}-DBM] sont terminées".format(self.codeSituation))
            # self.master.wait_window(myProgressBar)
            myProgressBar.close()
            json.dump({"codeSituation": self.codeSituation, "nbRowParsitua": self.nbRowParsitua, "listCode": self.listCode}, open("../var/cache/paramsDuplication.json", "w"))

            self.hadDuplicate = True
        except MyException as e:
            tkinter.messagebox.showerror(e.Title, e.Text)
            return

        return

    def duplicationSituationsMultipleAction(self):
        try:
            try:
                self.windowParamGestionSituation = Application().connect(title_re="Paramètres de la gestion des situations")
            except ElementNotFoundError as exc:
                raise MyException("Fenêtre inexistante", "La fenêtre des gestions des situations n'est pas ouverte")

            if (self.hadDuplicate):
                raise MyException("Données detectées", "ATTENTION : Il y a eu des duplications, il faut faire les affectations")

            # listCodeCustom = self.inputBoxCodeCustomGPGT()

            # if (not listCodeCustom): return

            self.duplicationSituationsAction(self.inputBoxCodeCustomGPGT())

        except MyException as e:
            tkinter.messagebox.showerror(e.Title, e.Text)
            return

    def chGroupeAction(self):
        _currentIndex = 0

        try:
            try:
                self.windowParamGestionSituation = Application().connect(title_re="Paramètres de la gestion des situations")
            except ElementNotFoundError as exc:
                raise MyException("Fenêtre inexistante", "La fenêtre des gestions des situations n'est pas ouverte")

            if (not self.codeSituation or len(self.listCode) == 0 or self.listCode[0] == "UN-"):
                raise MyException("Pas de données", "Il n'y a pas d'affectation à faire")

            self.checkCodeSituaStartMatrix()

            myProgressBar = ProgressBar(self, "Affectations", range(len(self.listCode) * self.nbRowParsitua), "Affectations du bloc [{0}-DBM] en cours...".format(self.codeSituation))
            myProgressBar.update()

            # NOTE: Début affectation
            for indexGroup, codeGroup in enumerate(self.listCode):
                for indexRow in range(self.nbRowParsitua):
                    _currentIndex = _currentIndex + 1

                    self.windowParamGestionSituation["Paramètres de la gestion des situations"]["TdxDBGrid1"].set_keyboard_focus()

                    if (indexRow == 0):
                        clipboard.copy(self.codeSituation + "-DBD")
                        self._sendKey("^v", 0, 0.323)
                        self._sendKey("{ENTER}", 0, 0.323)
                        self._sendKey("{RIGHT}", 0, 0.323)

                        clipboard.copy(codeGroup[1])
                        self._sendKey("^v", 0, 0.323)

                        self._sendKey("{ENTER}", 0, 0.323)
                        self._sendKey("{LEFT}", 0, 0.323)
                    elif (indexRow == self.nbRowParsitua - 1):
                        self._sendKey("{LEFT}", 0, 0.323)

                        clipboard.copy(self.codeSituation + "-FBD")
                        self._sendKey("^v", 0, 0.323)

                        self._sendKey("{ENTER}", 0, 0.323)
                    elif (self.copyCellContent() == "@C"):
                        clipboard.copy(codeGroup[0])
                        self._sendKey("^v", 0, 0.323)
                        self._sendKey("{ENTER}", 0, 0.323)

                    self.windowParamGestionSituation["Paramètres de la gestion des situations"]["Toutes"].type_keys("{SPACE}")

                    if (codeGroup[0] != "TOUTES"):
                        self.windowParamGestionSituation["Paramètres de la gestion des situations"]["Une seule"].type_keys("{SPACE}")

                        windowRechFicAnnex = self._waitExistWindow(winTitle="Recherche dans les Fichiers Annexes")

                        windowRechFicAnnex.top_window()
                        windowRechFicAnnex["Recherche dans les Fichiers Annexes"]["TEdit2"].set_keyboard_focus()
                        clipboard.copy(codeGroup[0])
                        self._sendKey("^v", 0.323, 0.323)
                        self._sendKey("{ENTER}")

                    self.windowParamGestionSituation["Paramètres de la gestion des situations"]["TdxDBGrid1"].set_keyboard_focus()
                    myProgressBar.num.set(_currentIndex)
                    myProgressBar.subTextVar.set("{0}/{1}".format(_currentIndex, myProgressBar.maximum))
                    myProgressBar.update()
                    self._sendKey("{DOWN}", 0.233, 0.323)
            # NOTE: Fin affectation

            self.listCode = []
            self.codeSituation = ""
            self.hadDuplicate = False
            remove("../var/cache/paramsDuplication.json")
            tkinter.messagebox.showinfo("Affectations CIES", "Affectations des codes terminées")
            myProgressBar.close()
        except MyException as e:
            tkinter.messagebox.showerror(e.Title, e.Text)
            return
        return
# #############################################################################################




# ################## Custom Function
    def listCodeGPGT(self, readExcel=True):
        if (self.codeSituation.startswith("GP-")):
            if (readExcel):
                self.listCode = self._readExcel(0, 1)
            else:
                for line in open('../data/Valeurs GP.csv'):
                    self.listCode.append(line.rstrip().replace("\"", "").split(";"))
        elif (self.codeSituation.startswith("GT-")):
            if (readExcel):
                self.listCode = self._readExcel(4, 5)
            else:
                for line in open('../data/Valeurs GTA.csv'):
                    self.listCode.append(line.rstrip().replace("\"", "").split(";"))
        elif (self.codeSituation.startswith("UN-")):
            self.listCode.append("UN-")
        else:
            self.codeSituation = None
            self.listCode = []

            raise MyException("Code situation", "Saisie incorrecte")

    def _readExcel(self, colStart, colEnd):
        listCode = []

        xlDataFrame = pd.read_excel("../data/Listing GTA pour Duplication-Affectation.xlsm", sheet_name="Valeurs")
        dataframe = xlDataFrame.iloc[:, colStart : colEnd + 1]

        for namedTuple in dataframe.itertuples(index=False):
            if (namedTuple[0] is not numpy.nan):
                listCode.append([namedTuple[0], namedTuple[1]])

        return listCode

    def sortedListCodeCustom(self, listCodeCustomGPGT, readExcel = True):
        _tmp = []

        if (len(self.listCode) == 1):
            if (len(listCodeCustomGPGT) > 1):
                self.listCode = []
                self.codeSituation = ""

                raise MyException("", "Il ne peut y avoir qu'un seul code pour une duplication unique")

            self.listCode = []

            if (readExcel):
                codesGP = self._readExcel(0, 1)
                codesGT = self._readExcel(4, 5)
                self.listCode = codesGP + codesGT
            else:
                for line in open("../data/Valeurs GP.csv"):
                    self.listCode.append(line.rstrip().replace("\"", "").split(";"))
                for line in open("../data/Valeurs GTA.csv"):
                    self.listCode.append(line.rstrip().replace("\"", "").split(";"))

        for customCode in listCodeCustomGPGT:
            _ = False

            for index2, code in enumerate(self.listCode):
                if (customCode[0] == code[0]):
                    _ = True

                    _tmp.insert(index2, code)
                    break

            if (_ == False):
                self.listCode = []
                self.codeSituation = ""
                raise MyException("", "Le code {0} n'existe pas dans le fichier CSV".format(customCode[0]))

        self.listCode = _tmp

        return self.listCode

    def getNbLigne(self):
        nbLigne = 0

        try :
            mydb = mysql.connector.connect(
                host=self.IPHost,
                port="3307",
                user="guest_alpha",
                password="alpha",
                database="alphaexpert"
            )
            cursor = mydb.cursor()

            cursor.execute(self.query_count_row, (self.codeSituation + "-DBM", self.codeSituation + "-FBM"))

            nbLigne = int(cursor.fetchone()[0])
        except Error as e:
            self.listCode = []
            self.codeSituation = ""
            raise MyException("MySQL Error", "Une erreur MySQL est survenue :\n\n{0}".format(e))
        finally:
            if (mydb.is_connected()):
                mydb.close()
                cursor.close()

        self.nbRowParsitua = nbLigne

        return nbLigne

    def copyCellContent(self, escapeBefore=False):
        clipboard.copy("")

        if (escapeBefore == True): self._sendKey("{ESC}", 0, 0.084)

        self._sendKey("{ENTER}", 0.23, 0)
        self._sendKey("^a^c" , 0.23, 0)
        self._sendKey("{ENTER}", 0.23, 0.23)

        return clipboard.paste()

    def checkCodeSituaStartMatrix(self):
        self.windowParamGestionSituation["Paramètres de la gestion des situations"]["TdxDBGrid1"].set_keyboard_focus()

        if (self.copyCellContent(True) != self.codeSituation + "-DBM"):
            raise MyException("Position incorrecte", "La cellule selectionnée ne correspond pas au code situation : {0}-DBM".format(self.codeSituation))

    def _sendKey(self, key="", sleepBefore=0.023, sleepAfter=0.023):
        sleep(sleepBefore)
        send_keys(key)
        sleep(sleepAfter)

    def _waitExistWindow(self, winTitle=""):
        window = None

        if (not winTitle):
            raise MyException("", "Invalid window title")

        while True:
            try:
                window = Application().connect(title_re=winTitle)
                break
            except Exception as e:
                sleep(1)

        return window

# ##################
# ################## GUI
    def inputBoxCodeSituation(self):
        try:
            self.RootInputBoxCodeSituation.destroy()
            self.RootInputBoxCodeSituation.quit()
        except Exception as e:
            pass
        try:
            self.RootInputBoxCodeCustomGPGT.destroy()
            self.RootInputBoxCodeCustomGPGT.quit()
        except Exception as e:
            pass

        width        = 233
        height       = 65

        self.RootInputBoxCodeSituation  = Tk()
        self.codeSituation              = StringVar(self.RootInputBoxCodeSituation)
        LabelCodeSituation              = Label(self.RootInputBoxCodeSituation, text="Code situation sans le -DBM")
        EntryCodeSituation              = Entry(self.RootInputBoxCodeSituation, textvariable=self.codeSituation   )
        ButtonOK                        = Button(self.RootInputBoxCodeSituation, text="OK", command=lambda: (self.RootInputBoxCodeSituation.destroy(), self.RootInputBoxCodeSituation.quit()))

        EntryCodeSituation.bind("<Return>", lambda event=None: (self.RootInputBoxCodeSituation.destroy(), self.RootInputBoxCodeSituation.quit()))
        EntryCodeSituation.bind("<Escape>", lambda event=None: (self.RootInputBoxCodeSituation.destroy()))

        LabelCodeSituation  .pack(fill=X)
        EntryCodeSituation  .pack(fill=X)
        ButtonOK            .pack(fill=X)
        EntryCodeSituation  .focus_set()

        self.RootInputBoxCodeSituation.resizable(False, False)
        self.RootInputBoxCodeSituation.wm_attributes("-topmost", 1)
        self.RootInputBoxCodeSituation.geometry('%dx%d+%d+%d' % (
            width, height,
            (self.RootInputBoxCodeSituation.winfo_screenwidth () / 2) - (width  / 2),
            (self.RootInputBoxCodeSituation.winfo_screenheight() / 2) - (height / 2)
        ))
        self.RootInputBoxCodeSituation.focus_force()
        self.RootInputBoxCodeSituation.mainloop() # Display InputBox

        self.codeSituation = self.codeSituation.get().upper()

        return self.codeSituation

    def inputBoxCodeCustomGPGT(self):
        listCodeCustomGPGT = []

        try:
            self.RootInputBoxCodeSituation.destroy()
            self.RootInputBoxCodeSituation.quit()
        except Exception as e:
            pass
        try:
            self.RootInputBoxCodeCustomGPGT.destroy()
            self.RootInputBoxCodeCustomGPGT.quit()
        except Exception as e:
            pass

        width        = 233
        height       = 65

        self.RootInputBoxCodeCustomGPGT = Tk()
        self.codeCustomGPGT             = StringVar(self.RootInputBoxCodeCustomGPGT)
        LabelCodeCustomGPGT             = Label(self.RootInputBoxCodeCustomGPGT, text="Saisir la liste des codes séparés par des \";\"")
        EntryCodeCustomGPGT             = Entry(self.RootInputBoxCodeCustomGPGT, textvariable=self.codeCustomGPGT                      )
        ButtonOK                        = Button(self.RootInputBoxCodeCustomGPGT, text="OK"
        , command=lambda: (self.RootInputBoxCodeCustomGPGT.destroy(), self.RootInputBoxCodeCustomGPGT.quit()))

        EntryCodeCustomGPGT.bind("<Return>", lambda event=None: (self.RootInputBoxCodeCustomGPGT.destroy(), self.RootInputBoxCodeCustomGPGT.quit()))
        EntryCodeCustomGPGT.bind("<Escape>", lambda event=None: (self.RootInputBoxCodeCustomGPGT.destroy()))

        LabelCodeCustomGPGT .pack(fill=X)
        EntryCodeCustomGPGT .pack(fill=X)
        ButtonOK            .pack(fill=X)
        EntryCodeCustomGPGT .focus_set()

        self.RootInputBoxCodeCustomGPGT.wm_attributes("-topmost", 1)
        self.RootInputBoxCodeCustomGPGT.geometry('%dx%d+%d+%d' % (
            width, height,
            (self.RootInputBoxCodeCustomGPGT.winfo_screenwidth() / 2) - (width / 2),
            (self.RootInputBoxCodeCustomGPGT.winfo_screenheight() / 2) - (height / 2)
        ))
        self.RootInputBoxCodeCustomGPGT.focus_force()
        self.RootInputBoxCodeCustomGPGT.mainloop()

        # Button OK clicked

        if (self.codeCustomGPGT.get() == ""):
            return False

        for index, code in enumerate(self.codeCustomGPGT.get().split(";")):
            if (code != ""):
                code = code.strip().upper()
                listCodeCustomGPGT.append([code, code])

        if (len(listCodeCustomGPGT) == 0):
            raise MyException("", "Il faut au moins 1 code")

        return listCodeCustomGPGT
# ################


def main():
    Root = Tk()
    app = Duplication_GS(Root)

    Root.mainloop()

    sys.exit(0)

def _json_object_hook(d): return namedtuple('X', d.keys())(*d.values())
def json2obj(data)      : return json.loads(data, object_hook=_json_object_hook)


if __name__ == "__main__":
    main()

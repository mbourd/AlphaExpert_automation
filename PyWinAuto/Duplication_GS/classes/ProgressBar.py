import tkinter.messagebox
from tkinter import IntVar, StringVar, Toplevel, X, simpledialog, ttk


class ProgressBar(simpledialog.Dialog):
    def __init__(self, parent, name="Untitled", lst=[], mainText="", subText=""):
        Toplevel.__init__(self, master=parent)

        self.titleprogress = name
        self.list          = lst
        self.length        = 400
        self.mainText      = mainText
        self.subText       = subText

        self.createWindow()
        self.createWidget()

    def createWindow(self):
        width = 400
        height = 70

        self.wm_attributes("-topmost", 1)

        self.protocol(u'WM_DELETE_WINDOW', self.close)
        self.resizable(False, False)
        # self.focus_set()
        # self.grab_self()
        self.transient(self.master)
        self.title(self.titleprogress)
        self.geometry("%dx%d+%d+%d" % (
            width, height,
            (self.winfo_screenwidth() / 2) - (width / 2),
            (self.winfo_screenheight() / 2) - (height / 2) + 233
        ))

    def createWidget(self):
        self.mainTextVar = StringVar()
        self.mainTextVar.set(self.mainText)

        self.subTextVar = StringVar()
        self.subTextVar.set(self.subText)

        self.num = IntVar()
        self.maximum = len(self.list)

        labelMainText = ttk.Label(self, textvariable=self.mainTextVar)
        progress = ttk.Progressbar(self,
            length=self.length,
            variable=self.num,
            mode="determinate",
            maximum=self.maximum
        )
        labelSubText = ttk.Label(self, textvariable=self.subTextVar)

        self.num.set(0)
        self.subTextVar.set("0/{0}".format(self.maximum))
        labelMainText.pack(fill=X)
        progress     .pack(fill=X)
        labelSubText.pack(fill=X)

    def close(self, event=None):
        self.destroy()

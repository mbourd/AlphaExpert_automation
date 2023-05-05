"""Python exceptions
"""

class MyException(Exception):
    Title = ""
    Text  = ""

    def __init__(self, Title="", Text=""):
        super(MyException, self).__init__()

        self.Title = Title
        self.Text  = Text

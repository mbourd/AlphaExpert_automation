
class First_():
    attr_First = "ok_first_"

    def __init__(self):
        self.okok = "okok"

        print("first")
        super().__init__()

    def testinherit(self):
        return "first test"
        pass
    pass

class First(First_):
    attr_First = "ok_first"

    def __init__(self):
        self.okok = "okokok"
        print("first__")
        super().__init__()

    def testinherit(self):
        return "first test"
        pass
    pass

class Second(First, First_):
    attr_Second = "ok_second"

    def __init__(self, var = ""):
        print("Second___")
        # print(var + " second")
        # super().__init__()
    def testinherit(self):
        self.attr_Outside = "out"
        print("second_test")
        return "second test"
        pass
    pass

class Outside2():
    attr_Outside = "ok_outside2"

    def __init__(self, var=""):
        print("init_outside2")
        super().__init__()

    def testinherit(self):
        print(self.attr_Outside)
        self.attr_Outside = "changed "
        return "outside2 test"
        pass
    pass
class Outside(Outside2, First):
    attr_Outside = "ok_outside"

    def __init__(self, var = ""):
        print("init outside")
        print(var + " outside")
        # super().__init__()

    def testinherit(self):
        print("outside test")
        return "outside test"
        pass
    pass

class Third(Second, Outside):
    attr_Outside = "dd"

    def __init__(self):
        self.attr_Outside = "lol"
        # super().__init__("ok")
        super(Outside, self).__init__("ok")
        # super(Third, self).testinherit()
        pass

    def test(self):
        # super().__init__()
        # print("2" + self.attr_Outside)
        # super(Outside, self).testinherit()
        # print(super(Outside, self).attr_Outside)
        # super().testinherit()
        # super(Outside, self).testinherit()
        # print(self.attr_First)
        # print(self.attr_Second)
        # print(self.attr_Outside)
        # print(super(Outside, self).testinherit())
        # print(super(First, self).testinherit())
        pass

    def testinherit(self):
        print("testinherit Third")
        pass


def main():
    first = Third()

    print("1" + first.attr_Outside)
    first.test()
    print("3" + first.attr_Outside)
    # first.testinherit()
    pass

if __name__ == "__main__":
    main()

def main():
    sets = {"ok", 1}
    sets.add("ok")
    print(sets.difference({1, 2, 3,"ko"}))
    print(sets.intersection({"ok"}))
    print(sets.symmetric_difference({"ok", 1}))
    print(sets.union({"ok"}, {"lol", "ok"}))

    dicts = {"ok": "lol"}
    print(dicts.get("dsqd", "sdd"))

    mlist = [1, 2, 3, 4]
    mlist.append(5)
    print(mlist)

    myTuple = (1, 2)
    print(myTuple.index(1))
    pass

if __name__ == "__main__":
    main()
    pass

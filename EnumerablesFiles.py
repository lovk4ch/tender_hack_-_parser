import os
import glob


def EnumerableDocx(PathDir):
    return glob.glob(PathDir + "/*.docx")


def EnumerableDoc(PathDir):
    return glob.glob(PathDir + "/*.doc")


def EnumerableOrder(SourseDir):
    return os.listdir(SourseDir)


def EnumerablePdf(PathDir):
    return glob.glob(PathDir + "/*.pdf")
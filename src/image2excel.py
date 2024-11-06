import os
import numpy.lib.recfunctions as nlr
import pandas as pd
from colormap import rgb2hex
from matplotlib.image import imread
from openpyxl import load_workbook
from PIL import Image

def __default_path(fileName):
    path = os.getcwd()
    file = os.path.join(path, fileName)
    return file

# oip9i9lijl

def __colorCells(string: str) -> str:
    return "background-color:" + string

def __changeZoom(excelFile, zoom = 10):
    wb = load_workbook(excelFile)
    for ws in wb.worksheets:
        ws.sheet_view.zoomScale = zoom
    wb.save(excelFile)

def __resize_picture(file, keep_aspect, extension):
    image = Image.open(file)
    if sorted(image.size)[0] > 260 or sorted(image.size)[1] > 300:
        if keep_aspect:
            image.thumbnail(
                tuple([300 / max(image.size) * x for x in image.size]), Image.ANTIALIAS
            )
        else:
            image = image.resize((260, 300), Image.ANTIALIAS)
        file = file.replace(extension, "_small" + extension)
        image.save(file)
    return file

def im2xlsx(file, resize = True, keep_aspect = False):
    path, fileName = os.path.split(file)
    extension = os.path.splitext(file)[1]
    if path == "":
        file = __default_path(fileName)
    if resize:
        file = __resize_picture(file, keep_aspect, extension)
    try:
        img = imread(file)
    except:
        raise Exception
    df = pd.DataFrame(nlr.unstructured_to_structured(img).astype("O"))
    df = df.applymap(lambda x: rgb2hex(*x))
    s = df.style.applymap(lambda x: "background-color:" + str(x))
    excelFile = file.replace(extension, ".xlsx")
    s.to_excel(excelFile, engine="xlsxwriter")
    __changeZoom(excelFile)

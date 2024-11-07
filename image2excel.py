import numpy.lib.recfunctions as nlr
import numpy
import pandas as pd
import os
from colormap import rgb2hex
from matplotlib.image import imread
from PIL import Image, ImageFile
from openpyxl import load_workbook
from typing import Union

#ignore missing imports for colormap -> mypy.ini ignore_missing_imports = True

def __default_path(fileName : str) -> str:
    path = os.getcwd()
    file = os.path.join(path,fileName)
    return file

def __colorCells(string : str) -> str:
    return 'background-color:' + string

def __changeZoom(excelFile : pd.io.excel._base.ExcelFile, zoom : int =10) -> None:
    wb = load_workbook(excelFile)
    for ws in wb.worksheets:
        ws.sheet_view.zoomScale = zoom
    wb.save(excelFile)

def __resize_picture(file, keep_aspect : bool, extension : str) -> Image.Image:
    image: Union[Image.Image, ImageFile.ImageFile]  = Image.open(file)
    if(sorted(image.size)[0]> 260 or sorted(image.size)[1]> 300):
        if keep_aspect:
            width, height = image.size
            new_size = (300 / max(image.size) * width, 300 / max(image.size) * height)
            image.thumbnail(new_size, Image.Resampling.LANCZOS)
            #image.thumbnail(tuple([300/max(image.size)*x for x in image.size]), Image.Resampling.LANCZOS) #Image.ANTIALIAS depricated
            #getting error cause the image can be over 2D
        else:
            image = image.resize((260,300), Image.Resampling.LANCZOS) #Image.ANTIALIAS depricated
        file=file.replace(extension,"_small" + extension)
        image.save(file)
    return file

def im2xlsx(file,  resize : bool =True, keep_aspect : bool =False) -> None:
    path, fileName = os.path.split(file)
    extension = os.path.splitext(file)[1]
    if (path == ""):
        file = __default_path(fileName)
    if resize:
        file = __resize_picture(file,keep_aspect,extension)
    try:
        img = imread(file)
    except:
        raise Exception
    df = pd.DataFrame(nlr.unstructured_to_structured(img).astype('O'))
    df=df.applymap(lambda x: rgb2hex(*x))
    s=df.style.applymap(lambda x:"background-color:"+str(x))
    excelFile = file.replace(extension,".xlsx")
    s.to_excel(excelFile,engine='xlsxwriter')
    __changeZoom(excelFile)
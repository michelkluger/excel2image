<h1>Documentation for image2excel</h1>

This library recieves as inputs an image and outputs an Excel file (.xlsx) where each cell will be a pixel of the image.

it is recommended to import the package in the following way:

    from image2excel import im2xlsx

The function has two parameters, resize (True is default) and keep_aspect (False is default), as shown by the following example:

    im2xlsx(file,resize=True,keepAspect=False)

The function has no return

---

The file parameter should contain a path, otherwise path where the script is running is taken as default path. The input path will be used for the output path as well.

The resize parameter will only work to reduce the image, in case the image is already smaller in both dimensions than the recommended 260 x 300, no resizing will take place

the parameter keep aspect, as it is relatively clear, will keep the ratio between width and height of the image 

---

This package and documentation are in the first version and it is also my first package in PYPI, so in case you have any sort of recommendations, please let me know =))

Also if you know how to make the code more performant or robust, I will also be happy with improving this 
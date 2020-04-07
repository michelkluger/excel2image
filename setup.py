from setuptools import setup

with open("README.md","r") as f:
    long_description = f.read()

setup(
    long_description=long_description,
    long_description_content_type="text/markdown",
    name='image2excel',
    version="0.0.4",
    description="converts images to excel tables",
    author_email='michel.kluger@gmail.com',
    license='MIT',
    py_modules=["image2excel"],
    install_requires=[
          'pandas',
          'Pillow',
          'colormap',
          'openpyxl',
          'XlsxWriter',
          'Jinja2',
          'easydev'
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3',
    package_dir={'':'src'}
    )
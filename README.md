# Excel-to-JSON
This is the source code for an excel to json convert with an executable file.

Go to the [dist](./dist) folder to get the executable file!

# Libraries used
## Python libraries
Imported default python libraries used
~~~
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import json
import os
~~~
## Pip libraries
### Pandas
Install pandas:
~~~
pip install pandas openpyxl
~~~
Import pandas:
~~~
import pandas as pd
~~~
### Pyinstsaller
Used pyinstaller to convert to .exe file

Install pyinstaller:
~~~
pip install pyinstaller
~~~
Convert file to .exe without PATH being updated:
~~~
python -m PyInstaller --onefile --icon=convert-icon.ico --noconsole excel-reader.py
~~~

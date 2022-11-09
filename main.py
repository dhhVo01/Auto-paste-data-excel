import pyautogui as pg
import time
import pyperclip
import pandas as pd
import numpy as np
import statistics
from statistics import mode

path = 'C:/Users/vohuy/Desktop/divide _data.xlsx'
sheets = ["Khánh", "Vĩ", "Nhi"]
skip_rows = [i for i in range(1, 2001)]
n_rows = 1000
use_cols = "A:G"
df = pd.read_excel(path, sheet_name = sheets, skiprows = skip_rows, nrows = n_rows, usecols = use_cols)
cols = df[sheets[0]].columns
for sheet in sheets:
    df[sheet] = df[sheet].replace(to_replace = np.nan, value =0)
def auto_tag(value):
    pyperclip.copy(value)
    pg.hotkey("ctrl", "v")
    pg.press("right")

def create_data_tag(df):
    n_cols = cols[1:].size
    n_sheets = len(sheets)
    for key in range(n_rows):
        for i in range(n_cols):
            value = (int)(df[sheets[i%n_sheets]][cols[i+1]][key])
            auto_tag(value)
        pg.press("down")
        pg.press("left", presses=6)
        time.sleep(3)
        sheets.append(sheets[0])
        sheets.remove(sheets[0])
time.sleep(10)
create_data_tag(df)

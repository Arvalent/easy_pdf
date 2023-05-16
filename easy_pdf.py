import os
import gc
from tkinter import *
from inspect import getsourcefile
from os.path import abspath
from win32api import GetSystemMetrics
width = GetSystemMetrics (0)
height = GetSystemMetrics (1)

path = abspath(getsourcefile(lambda:0))
path = os.path.split(path)[0]
os.chdir(path)

from interface import *
from default_vars import *


window = Tk()
window.iconbitmap("logo.ico")
window.title('Easy PDF')
window.geometry("%dx%d%+d%+d" % (window_width, window_height, (width - window_width) // 2, (height - window_height) // 2))
window.configure(bg=window_background_color)

interface = Interface(window)

window.mainloop()

del window
gc.collect()
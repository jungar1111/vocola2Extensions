#as of the writing of these Vocola extensions, Vocola is not taking more than 1 variable per
#(def) functions
#that's why I'm using global variables
#test

import sys
import pywinauto
from pywinauto import application
from pyautogui import getWindowsWithTitle
import tkinter as tk
from win32gui import GetWindowText, GetForegroundWindow

def commonvars():
    global hwnd
    global dlg_spec
    global currxposval
    global curryposval
    global currwidth
    global currheight
    global dlg_spec
    global app

    #setting up for move/resize
    Title = GetWindowText(GetForegroundWindow())
        
    app = application.Application(backend="win32").connect(title_re=Title)
    app.top_window().set_focus()
        
    hwnd = getWindowsWithTitle(Title)
    dlg_spec = app.window(title=Title)
    currxposval = int(hwnd[0].left)
    curryposval = int(hwnd[0].top)
    currwidth = int(hwnd[0].width)
    currheight = int(hwnd[0].height)    
        
# Vocola function: placewindow.lengthen
def lengthen():
    global dlg_spec
    global app
    root = tk.Tk()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    additionalHeight = 120
    
    commonvars()
    
    newheight = currheight + additionalHeight
    
    try:
        app.top_window().set_focus()
        dlg_spec.move_window(currxposval, curryposval, currwidth, newheight, repaint=True)
        app.top_window().set_focus()
    #except ElementAmbiguousError:
    #    print("Too many of this type window")
    #    for i in range(1, len(dlg_spec)):
    #        print(i,": ", dlg_spec,end = " ")
    #except ElementNotFoundError:
    #    print("Window not found")
    except FileNotFoundError:
        print("Window with that title not running")

# Vocola function: placewindow.shorten
def shorten():
    global dlg_spec
    global app
    root = tk.Tk()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    additionalHeight = -120
    
    commonvars()
    
    newheight = currheight + additionalHeight
    print("new height - \n",newheight) 
    print("currheight - \n",currheight) 
    print("additionalHeight - \n",additionalHeight)
    
    try:
        app.top_window().set_focus()
        dlg_spec.move_window(currxposval, curryposval, currwidth, newheight, repaint=True)
        app.top_window().set_focus()
    #except ElementAmbiguousError:
    #    print("Too many of this type window")
    #    for i in range(1, len(dlg_spec)):
    #        print(i,": ", dlg_spec,end = " ")
    #except ElementNotFoundError:
    #    print("Window not found")
    except FileNotFoundError:
        print("Window with that title not running")

# Vocola function: placewindow.widen
def widen():
    global dlg_spec
    global app
    root = tk.Tk()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    additionalWidth = 120
 
    commonvars()
    
    newwidth = currwidth + additionalWidth
        
    try:
        app.top_window().set_focus()    
        dlg_spec.move_window(currxposval, curryposval, newwidth, currheight, repaint=True)
        app.top_window().set_focus()
    #except ElementAmbiguousError:
    #    print("Too many of this type window")
    #    for i in range(1, len(dlg_spec)):
    #        print(i,": ", dlg_spec,end = " ")
    #except ElementNotFoundError:
    #    print("Window not found")
    except FileNotFoundError:
        print("Window with that title not running")
        
# Vocola function: placewindow.narrow
def narrow():
    global dlg_spec
    global app
    root = tk.Tk()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    additionalWidth = -120
 
    commonvars()
    
    newwidth = currwidth + additionalWidth
        
    try:
        app.top_window().set_focus()    
        dlg_spec.move_window(currxposval, curryposval, newwidth, currheight, repaint=True)
        app.top_window().set_focus()
    #except ElementAmbiguousError:
    #    print("Too many of this type window")
    #    for i in range(1, len(dlg_spec)):
    #        print(i,": ", dlg_spec,end = " ")
    #except ElementNotFoundError:
    #    print("Window not found")
    except FileNotFoundError:
        print("Window with that title not running")
        
# Vocola function: placewindow.upperleft
def upperleft():
    global dlg_spec
    global app
    root = tk.Tk()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
        
    commonvars()
    
    try:
        app.top_window().set_focus()
        dlg_spec.move_window(0, 0, currwidth, currheight, repaint=True)
        app.top_window().set_focus()
    #except ElementAmbiguousError:
    #    print("Too many of this type window")
    #    for i in range(1, len(dlg_spec)):
    #        print(i,": ", dlg_spec,end = " ")
    #except ElementNotFoundError:
    #    print("Window not found")
    except FileNotFoundError:
        print("Window with that title not running")

# Vocola function: placewindow.upperright
def upperright():
    global dlg_spec
    global app
    root = tk.Tk()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    commonvars()
        
    try:
        app.top_window().set_focus()
        dlg_spec.move_window((screen_width-currwidth), 0, currwidth, currheight, repaint=True)
        app.top_window().set_focus()
    #except ElementAmbiguousError:
    #    print("Too many of this type window")
    #    for i in range(1, len(dlg_spec)):
    #        print(i,": ", dlg_spec,end = " ")
    #except ElementNotFoundError:
    #    print("Window not found")
    except FileNotFoundError:
        print("Window with that title not running")
        
# Vocola function: placewindow.lowerright
def lowerright():
    global dlg_spec
    global app
    root = tk.Tk()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    commonvars()
    
    try:
        app.top_window().set_focus()
        dlg_spec.move_window((screen_width-currwidth), (screen_height-currheight),currwidth, currheight, repaint=True)        
        dlg_spec.move_window((screen_width-currwidth), (screen_height-currheight), 0, currwidth, currheight, repaint=True)
        app.top_window().set_focus()
    #except ElementAmbiguousError:
    #    print("Too many of this type window")
    #    for i in range(1, len(dlg_spec)):
    #        print(i,": ", dlg_spec,end = " ")
    #except ElementNotFoundError:
    #    print("Window not found")
    except FileNotFoundError:
        print("Window with that title not running")
        
# Vocola function: placewindow.lowerleft
def lowerleft():
    global dlg_spec
    global app
    root = tk.Tk()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    commonvars()
    
    try:
        app.top_window().set_focus()
        dlg_spec.move_window((currwidth-screen_width), (screen_height-currheight),currwidth, currheight, repaint=True)        
        dlg_spec.move_window(0, (screen_height-currheight), currwidth, currheight, repaint=True)
        app.top_window().set_focus()
    #except ElementAmbiguousError:
    #    print("Too many of this type window")
    #    for i in range(1, len(dlg_spec)):
    #        print(i,": ", dlg_spec,end = " ")
    #except ElementNotFoundError:
    #    print("Window not found")
    except FileNotFoundError:
        print("Window with that title not running")
        
# Vocola function: placewindow.center
def center():
    global dlg_spec
    global app
    root = tk.Tk()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    commonvars()
    
    try:
        app.top_window().set_focus()      
              
        dlg_spec.move_window(
            int(round((screen_width/2)-(currwidth/2))), 
            int(round((screen_height/2)-(currheight/2))),
            currwidth, currheight, repaint=True)
        app.top_window().set_focus()
    #except ElementAmbiguousError:
    #    print("Too many of this type window")
    #    for i in range(1, len(dlg_spec)):
    #        print(i,": ", dlg_spec,end = " ")
    #except ElementNotFoundError:
    #    print("Window not found")
    except FileNotFoundError:
        print("Window with that title not running")
    

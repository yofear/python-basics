# python-basics
# dictionary attack script

from xlrd import *
import sys
import win32com.client
import os

xlApp = win32com.client.Dispatch("Excel.Application")
filename = sys.argv[1]
path = os.getcwd()
try:
	wb = xlApp.Workbooks.Open(path+'\\'+filename, false, true, none, "123")
	print("Success! Password is: 123")
except:
	print("Incorrect password")
	pass

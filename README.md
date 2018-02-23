# python-basics
dictionary attack script

from xlrd import *
import sys
import win32com.client
import os

xlApp = win32com.client.Dispatch("Excel.Application")
filename = sys.argv[1]
path = os.getcwd()

password_file = open ( 'wordlist.txt', 'r' )
passwords = password_file.readlines()
password_file.close()

passwords = [item.rstrip('\n') for item in passwords]

results = open('results.txt', 'w')

for password in passwords:
	print(password)
	try:
		wb = xlApp.Workbooks.Open(path+'\\'+filename, false, true, none, password)
		print("Success! Password is: "+password)
		results.write(password)
		results.close()
	except:
		print("Incorrect password")
		pass

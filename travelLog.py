#Imports all the needed modules to make this program work.
import openpyxl
from openpyxl import workbook
from openpyxl.styles import Font
from time import strftime
from tkinter import *
from travelData import *

#Creates variables containing the current date, current weekday and current month.
date = strftime('%d.%m.%Y')
weekday = strftime('%A')
month = strftime ('%B')

#Loads in the Excel workbook and selects this month's sheet.
#If this month's sheet doesn't exist, it will create it and write in the headers of each column.
wb = openpyxl.load_workbook(excelWorkBook)
if month in wb.sheetnames:
	ws = wb[month]
else:
	wb.create_sheet(month)
	ws = wb[month]
	ws['A1'] = 'Weekday'
	ws['A1'].font = Font(bold=True)
	ws['B1'] = 'Date'
	ws['B1'].font = Font(bold=True)
	ws['C1'] = 'Workplace'
	ws['C1'].font = Font(bold=True)
	ws['D1'] = 'Travel route'
	ws['D1'].font = Font(bold=True)
	ws['E1'] = 'Distance'
	ws['E1'].font = Font(bold=True)
	ws['F1'] = 'Comment'
	ws['F1'].font = Font(bold=True)

#Defining the GUI.
root = Tk()
root.title('Travel Log')
carIcon = PhotoImage(file='car.png')
root.tk.call('wm', 'iconphoto', root._w, carIcon)

#Creates variables to be filled in by user input and button presses.
travelRoute = ''
totalDist = 0
travelSteps = []
workPlace = ''
textInput = ''

#Defining functions for the different buttons in the program.

#This button is for adding the home address to the variable travelRoute.
def homeButton():
    global travelSteps
    global travelRoute
    travelSteps = travelSteps + [0]
    if travelRoute == '':
        travelRoute = travelRoute + Home
    else:
        travelRoute = travelRoute + ' - ' + Home

#This button is for adding the work1 address to the variable travelRoute.
def work1Button():
    global travelSteps
    global travelRoute
    travelSteps = travelSteps + [1]
    if travelRoute == '':
        travelRoute = travelRoute + Work1
    else:
        travelRoute = travelRoute + ' - ' + Work1

#This button is for adding the work2 address to the variable travelRoute.
def work2Button():
    global travelSteps
    global travelRoute
    travelSteps = travelSteps + [2]
    if travelRoute == '':
        travelRoute = travelRoute + Work2
    else:
        travelRoute = travelRoute + ' - ' + Work2

#This button is for adding the work3 address to the variable travelRoute.
def work3Button():
    global travelSteps
    global travelRoute
    travelSteps = travelSteps + [3]
    if travelRoute == '':
        travelRoute = travelRoute + Work3
    else:
        travelRoute = travelRoute + ' - ' + Work3

#This button is for adding the work4 address to the variable travelRoute.
def work4Button():
    global travelSteps
    global travelRoute
    travelSteps = travelSteps + [4]
    if travelRoute == '':
        travelRoute = travelRoute + Work4
    else:
        travelRoute = travelRoute + ' - ' + Work4

#This button is for adding the work5 address to the variable travelRoute.
def work5Button():
    global travelSteps
    global travelRoute
    travelSteps = travelSteps + [5]
    if travelRoute == '':
        travelRoute = travelRoute + Work5
    else:
        travelRoute = travelRoute + ' - ' + Work5

#This button is for storing the text written into the text box into the variable textInput.
def retrieveInput():
    global textInput
    textInput = textInputButton.get()

#This function makes it so that pressing the enter key triggers the above function.
def enterPress(event):
    retrieveInput()

#This binds any press on the enter key to the function above.
root.bind('<Return>', enterPress)

#This function determines the total travel distance, and which workplace(s) you've worked at.
#Then it appends it to the next row in this month's sheet in your Excel workbook.
def writeToExcel():
	global travelDist
	global travelSteps
	global travelRoute
	global travelSteps
	global textInput
	global workPlace
	global ws

	travelSteps.sort()

	#These statements determine the total travel distance by looking at which locations you've visited.
	if travelSteps == [0, 0, 1]:
		travelDist = distHomeWork1 + distHomeWork1
	if travelSteps == [0, 0, 2]:
		travelDist = distHomeWork2 + distHomeWork2
	if travelSteps == [0, 0, 3]:
		travelDist = distHomeWork3 + distHomeWork3
	if travelSteps == [0, 0, 4]:
		travelDist = distHomeWork4 + distHomeWork4
	if travelSteps == [0, 0, 5]:
		travelDist = distHomeWork5 + distHomeWork5
	if travelSteps == [0, 0, 1, 2]:
		travelDist = distHomeWork1 + distWork1Work2 + distHomeWork2
	if travelSteps == [0, 0, 1, 3]:
		travelDist = distHomeWork1 + distWork1Work3 + distHomeWork3
	if travelSteps == [0, 0, 1, 4]:
		travelDist = distHomeWork1 + distWork1Work4 + distHomeWork4
	if travelSteps == [0, 0, 1, 5]:
		travelDist = distHomeWork1 + distWork1Work5 + distHomeWork5
	if travelSteps == [0, 0, 2, 3]:
		travelDist = distHomeWork2 + distWork2Work3 + distHomeWork3
	if travelSteps == [0, 0, 2, 4]:
		travelDist = distHomeWork2 + distWork2Work4 + distHomeWork4
	if travelSteps == [0, 0, 2, 5]:
		travelDist = distHomeWork2 + distWork2Work5 + distHomeWork5
	if travelSteps == [0, 0, 3, 4]:
		travelDist = distHomeWork3 + distWork3Work4 + distHomeWork4
	if travelSteps == [0, 0, 3, 5]:
		travelDist = distHomeWork3 + distWork3Work5 + distHomeWork5
	if travelSteps == [0, 0, 4, 5]:
		travelDist = distHomeWork4 + distWork4Work5 + distHomeWork5
	else:
		pass

	#These statements determines which workplace you worked at by calculating the total distance traveled.
	if travelDist == distHomeWork1 + distHomeWork1:
		workPlace = wpWork1
	if travelDist == distHomeWork2 + distHomeWork2:
		workPlace = wpWork2
	if travelDist == distHomeWork3 + distHomeWork3:
		workPlace = wpWork3
	if travelDist == distHomeWork4 + distHomeWork4:
		workPlace = wpWork4
	if travelDist == distHomeWork5 + distHomeWork5:
		workPlace = wpWork5
	if travelDist == distHomeWork1 + distWork1Work2 + distHomeWork2:
		workPlace = wpWork1 + ' og ' + wpWork2
	if travelDist == distHomeWork1 + distWork1Work3 + distHomeWork3:
		workPlace = wpWork1 + ' og ' + wpWork3
	if travelDist == distHomeWork1 + distWork1Work4 + distHomeWork4:
		workPlace = wpWork1 + ' og ' + wpWork4
	if travelDist == distHomeWork1 + distWork1Work5 + distHomeWork5:
		workPlace = wpWork1 + ' og ' + wpWork5
	if travelDist == distHomeWork2 + distWork2Work3 + distHomeWork3:
		workPlace = wpWork2 + ' og ' + wpWork3
	if travelDist == distHomeWork2 + distWork2Work4 + distHomeWork4:
		workPlace = wpWork2 + ' og ' + wpWork4
	if travelDist == distHomeWork2 + distWork2Work5 + distHomeWork5:
		workPlace = wpWork2 + ' og ' + wpWork5
	if travelDist == distHomeWork3 + distWork3Work4 + distHomeWork4:
		workPlace = wpWork3 + ' og ' + wpWork4
	if travelDist == distHomeWork3 + distWork3Work5 + distHomeWork5:
		workPlace = wpWork3 + ' og ' + wpWork5
	if travelDist == distHomeWork4 + distWork4Work5 + distHomeWork5:
		workPlace = wpWork4 + ' og ' + wpWork5
	else:
		pass

	#This appends the collected data into the next row in this month's sheet in your Excel workbook.
	ws.append([weekday, date, workPlace, travelRoute, travelDist, textInput])

	#Empty all variables before a possible new append.
	travelRoute = ''
	totalDist = 0
	travelSteps = []
	textInput = ''
	workPlace = ''

	#Writes all the data to the Excel workbook.
	wb.save(excelWorkBook)

#This function does the same as the above, but also closes the app window.
def writeToExcelClose():
	writeToExcel()
	root.destroy()

#Buttons in the GUI with button titles, sizes and grid definitions.
#Change the text on the buttons b1, b2, b3, b4, b5 and b6 to your own values here.
b1 = Button(root, text='Home', command=homeButton, width=13)
b1.grid(row=0, column=0)
b2 = Button(root, text=wpWork1, command=work1Button, width=13)
b2.grid(row=0, column=1)
b3 = Button(root, text=wpWork2, command=work2Button, width=13)
b3.grid(row=0, column=2)
b4 = Button(root, text=wpWork3, command=work3Button, width=13)
b4.grid(row=1, column=0)
b5 = Button(root, text=wpWork4, command=work4Button, width=13)
b5.grid(row=1, column=1)
b6 = Button(root, text=wpWork5, command=work5Button, width=13)
b6.grid(row=1, column=2)
commentLabel = Label(root, text='Comment here:')
commentLabel.grid(row=2, column=0, pady=5)
textInputButton = Entry(root, bd=1)
textInputButton.grid(row=2, column=1, pady=5)
b7 = Button(root, text='Submit comment', command=retrieveInput, width=13)
b7.grid(row=2, column=2, pady=5)
b8 = Button(root, text='Save and new', command=writeToExcel, width=13)
b8.grid(row=3, column=0, columnspan=2)
b9 = Button(root, text='Save and close', command=writeToExcelClose, width=13)
b9.grid(row=3, column=1, columnspan=2)

#Give focus to the text input area.
textInputButton.focus()

#This keeps the program running until the windows is crossed out by pressing the 'X' button.
root.mainloop()

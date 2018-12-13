# TravelLog

A simple program written in Python 3.x that lets you log your driving to and from your work places in a spreadsheet.

## Getting started

Clone the files to your computer and edit travelData.py with the addresses, locations and distances to fit your most commonly visited locations.
After you've added your own data to travelData.py, just run travelLog.py to start using it.

### Prerequisites

To run this program you need to pip install OpenPyXL

```
pip install openpyxl
```

### What does it do?

To use this program all you have to do is click on the locations you've visited that day i.e. Home -> Location 1 -> Location 2 -> Home. In this version only 2 locations in addition to home is supported.

![GUI](https://github.com/Myrheimb/TravelLog/blob/master/Images/GUI.png)

When you click save your travel route will be saved to the Excel file specified with the following headers and information.
A new sheet will be made for every new month automatically.

![Excel](https://github.com/Myrheimb/TravelLog/blob/master/Images/Example.png)

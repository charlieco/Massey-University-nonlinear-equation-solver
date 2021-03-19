##importexcelsheets():
Loads the Excel file to read the available sheeps ready for selection in the UI 
Returns: nothing
Arguments: None 

##importexcel(sheet):
Loads the Excel file and user selected Excel sheet Ok then import the three columns from the sheet into globe lists. The end of the function automatically calls the trans function. 
Returns: nothing
Arguments: sheet ( holds the Excel sheet selected by the user in the interface)

##exportexcel(sheet, tosave):
Saves the variables and solutions back to the same Excel sheet. 
Returns: nothing
Arguments: sheet( holds the Excel sheet selected by the user in the interface), tosave(Holds the solutions for the equations)

# function overview

## importexcelsheets():
Loads the Excel file to read the available sheeps ready for selection in the UI 
Returns: nothing
Arguments: None 

## importexcel(sheet):
Loads the Excel file and user selected Excel sheet Ok then import the three columns from the sheet into globe lists. The end of the function automatically calls the trans function. 
Returns: nothing
Arguments: sheet ( holds the Excel sheet selected by the user in the interface)

## exportexcel(sheet, tosave):
Saves the variables and solutions back to the same Excel sheet. 
Returns: nothing
Arguments: sheet( holds the Excel sheet selected by the user in the interface), tosave(Holds the solutions for the equations)

## trans():
translation function for translating the ees equations to a format that is that can be used in python. Automatically runs the limits function at the end. 
Returns: nothing
Arguments: None 
the initial equations for translation are stored in the variable working_list. The equations come in two forms. 
1. The first type of equation only contains: variables, numbers and standard math symbols. 
2. the second type of Equation contains a function call. Which needs to be in the following form answer=functionname(‘Coolant’,a=ate,b=bee) At then gets translated into the following form ['functionname_ab', ‘Coolant’, a, b, answer].  the following is a sample of before and after h1=enthalpy('R744',P=p1,T=t1), ['enthalpy_pt', 'R744', p1, t1, h1]. The coolant can be available if you needed. If the coolant is not needed in the function It can just be replaced with ‘na’. 

## lims():
Calculates the thermodynamic limits So that the cool pop functions don't get the values that they can't accept. 
Returns: nothing
Arguments: None 


## con(x, flist, vlist, na):
Evaluates the list of Equations by substituting values into the equations. Takes the absolute of all output values converts the roots problem into a minimisation one. 
Returns: A list of answers to the evaluated equations
Arguments: x(List of values for each variable), flist(List of equations), vlist(List of variables), na(not used)

## con_single
Evaluates a single equation  by substituting values into the equations. 
Returns: A single number answer 
Arguments: x(List of values for each variable), flist(single equations), vlist(List of variables), na(not used)


## Conj
Evaluates The jacobian Matrix by substituting values into the equations.
Returns: A list of answers to the evaluated equations
Arguments: x(List of values for each variable), flist(List of equations), vlist(List of variables), j(The jacobian Matrix)

## Conj_lam(vlist, j, x):
For evaluating the matrix in parallel
Returns: A single row of the matrix
Arguments: x(List of values for each variable), vlist(List of variables), j(1 row of the jacobian Matrix)


# Function & libraries

## libraries Not included in python base package

### sympy
The library for adding in symbolic equations
<br/>https://www.sympy.org/en/index.html

### openpyxl
openpyxl is a Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files.
<br/>https://openpyxl.readthedocs.io/en/stable/

### CoolProp
CoolProp is the thermophysical property database used
<br/>http://www.coolprop.org/

### numpy
Adds additional mathematical  functionality 
<br/>https://numpy.org/


### scipy.optimize
SciPy optimize provides functions for minimizing (or maximizing) objective functions
<br/>https://docs.scipy.org/doc/scipy/reference/optimize.html

### fluprodia
Improved cool prop diagram drawing library 
<br/>https://pypi.org/project/fluprodia/

### PySimpleGUI
Basic Python gui library
<br/>https://pysimplegui.readthedocs.io/en/latest/

### joblib
Joblib is a set of tools to provide lightweight pipelining in Python. used for simple parallel computing.
<br/>https://joblib.readthedocs.io/en/latest/

## Functions

### importexcelsheets():
Loads the Excel file to read the available sheeps ready for selection in the UI 
<br/>Returns: nothing
<br/>Arguments: None 

### importexcel(sheet):
Loads the Excel file and user selected Excel sheet Ok then import the three columns from the sheet into globe lists. The end of the function automatically calls the trans function. 
<br/>Returns: nothing
<br/>Arguments: sheet ( holds the Excel sheet selected by the user in the interface)

### exportexcel(sheet, tosave):
Saves the variables and solutions back to the same Excel sheet. 
<br/>Returns: nothing
<br/>Arguments: sheet( holds the Excel sheet selected by the user in the interface), tosave(Holds the solutions for the equations)

### trans():
translation function for translating the ees equations to a format that is that can be used in python. Automatically runs the limits function at the end. 
<br/>Returns: nothing
<br/>Arguments: None 
<br/>the initial equations for translation are stored in the variable working_list. The equations come in two forms. 
1. The first type of equation only contains: variables, numbers and standard math symbols. 
2. the second type of Equation contains a function call. Which needs to be in the following form answer=functionname(‘Coolant’,a=ate,b=bee) At then gets translated into the following form ['functionname_ab', ‘Coolant’, a, b, answer].  the following is a sample of before and after h1=enthalpy('R744',P=p1,T=t1), ['enthalpy_pt', 'R744', p1, t1, h1]. The coolant can be available if you needed. If the coolant is not needed in the function It can just be replaced with ‘na’. 


### lims():
Calculates the thermodynamic limits So that the cool pop functions don't get the values that they can't accept. 
<br/>Returns: nothing
<br/>Arguments: None 


### con(x, flist, vlist, na):
Evaluates the list of Equations by substituting values into the equations. Takes the absolute of all output values converts the roots problem into a minimisation one. 
<br/>Returns: A list of answers to the evaluated equations
<br/>Arguments: x(List of values for each variable), flist(List of equations), vlist(List of variables), na(not used)

### con_single
Evaluates a single equation  by substituting values into the equations. 
<br/>Returns: A single number answer 
<br/>Arguments: x(List of values for each variable), flist(single equations), vlist(List of variables), na(not used)


### Conj
Evaluates The jacobian Matrix by substituting values into the equations.
<br/>Returns: A list of answers to the evaluated equations
<br/>Arguments: x(List of values for each variable), flist(List of equations), vlist(List of variables), j(The jacobian Matrix)

### Conj_lam(vlist, j, x):
For evaluating the matrix in parallel
<br/>Returns: A single row of the matrix
<br/>Arguments: x(List of values for each variable), vlist(List of variables), j(1 row of the jacobian Matrix)

### Algebra_solver
Attempts to solve each equation algebraically. This is done by initially looking for equations that only have one unsolved variable. If the equation does not feature a function The program uses the sympy function solve which can rearrange for the unsolved variable Then we can just substitute in all the known variables and get the answer.  For equations that do feature a function it initially calculates the solution with a randomly calculated guess. Since we know this equation equals 0 we know how far away the solution is from the correct answer.  we then  try taking away and adding that to the guess.  this should result in the answer to this variable. 
<br/>Returns: nothing
<br/>Arguments: None 

### Jacobian
jacobian matrix is a two dimensional Matrix full of the gradients Each row is the gradient for every variable in the system for one equation
<br/>Returns: the Jacobian Matrix 
<br/>Arguments: None 

### Non_linear_solver
Runs the main nonlinear Solver from scipy.  The solver is from the optimise package and uses the least squares algorithm. This function also generates the limits based on the letters used in the variable name. 
Link to sicpy function https://docs.scipy.org/doc/scipy/reference/generated/scipy.optimize.least_squares.html
<br/>Returns: sol_out(list of variables and the solutions)
<br/>Arguments: tol(Stores the tolerance that the Solver runs to)

### Diagram_draw
For drawing the final diagram 
<br/>Returns: nothing
<br/>Arguments: None 

### solver
the main Solver function that is run when the UI button is clicked. 
<br/>Returns: nothing
<br/>Arguments: None 

# Polynomial Fitting

This is just a simple tool to help fit data to a polynomial of your choosing. It consistes of five modules, four with the **mod** prefix to denote separate libraries of functions and the fifth with the **prg** prefix to denote the main program where we will generate the polynomial fits.

## Getting Started

Let's now go over some of the basic steps to gain familiarity. We will also highlight some extra useful features later on. 

First, we'll start with choosing a **.csv** file. This will be the only file type supported, so prepare your data accordingly. It's important to note that you are not restricted to importing only two columns of data, nor any specific columnar formatting. However, this does not support headers, so the data should be numeric only.

```VBA
'choose the .csv file you'd like
csvFilepath = modText.csvFind
```

We'll then parse the file and load it into an array in memory.

```VBA
csvMatrix = modText.csvParse(csvFilepath)
```

As an optional step, assuming ***csvMatrix*** has more than two columns of data - (x,y), we can indicidually specify what our chosen x and y arrays will be. These individual arrays are dimensioned as Nx1, (rxc). So, for instance, if you had ***csvMatrix*** having three columns of data - (x, y1, y2), and we only wanted to perform a fit to (x,y2) then we would pull out individual vectors as such:

```VBA
'Separate x and y arrays from larger csvMatrix, if needed
dataX = modMatrix.matVec(csvMatrix, 1)
dataY = modMatrix.matVec(csvMatrix, 3)
```



## Extra Features

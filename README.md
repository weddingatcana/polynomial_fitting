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

As an optional step, assuming ***csvMatrix*** has more than two columns of data - (x,y), we can individually specify what our chosen x and y arrays will be. These individual arrays are dimensioned as Nx1, (rxc). So, for instance, if you had ***csvMatrix*** having three columns of data - (x, y1, y2), and we only wanted to perform a fit to (x,y2) then we would pull out individual vectors as such:

```VBA
'Separate x and y arrays from larger csvMatrix, if needed
dataX = modMatrix.matVec(csvMatrix, 1)
dataY = modMatrix.matVec(csvMatrix, 3)
```

We'd then recombine both vectors into a single 2D array with specific columnar formatting - (x,y2):

```VBA
data_2D = modMatrix.matJoin(dataX, dataY)
```

This is to staisfy the structural requirements imposed on input arrays given by the polynomial fitting function, ***optPolyFit***. Now, satisfied with our data formatting, we can specify the order of polynomial we'd like to fit to our data, ***data_2D***.

```VBA
data_Fit = modOptimization.optPolyFit(data_2D, 5)
```
The function above, ***optPolyFit*** has two inputs - the 2D data, and the order of polynomial to fit. For this example, we've chosen a fifth order polynomial. Once the function has completed, we now seek to export the calculated data, ***data_Fit***, and perhaps other fields. We'll define a boolean variable, ***csvStatus***, to display true/false if the exporting was successful. We can see this in practice below:

```VBA
'Write the various results to written filepaths
csvStatus = modText.csvWrite(csvMatrix, "raw.csv")
csvStatus = modText.csvWrite(data_2D, "xy.csv")
csvStatus = modText.csvWrite(data_Fit, "order5.csv")
```

We see that the function ***csvWrite*** takes two inputs as well - single or multidimensional data, and the filename with extension. The only extension one should use is, as you guessed, **.csv**. An optional third argument is also provided by ***csvWrite*** which allows one to provide a directory path; however, by default one is provided for export to the desktop. 

## Extra Features

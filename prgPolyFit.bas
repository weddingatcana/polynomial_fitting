Attribute VB_Name = "prgPolyFit"
Option Explicit

Sub PolyFit()

    Dim csvStatus As Boolean, _
        csvFilepath$, _
        csvMatrix#(), _
 _
        dataX#(), _
        dataX_dS#(), _
        dataY#(), _
        dataY_dS#(), _
        data_dS#(), _
        data_2D#(), _
        data_Fit#(), _
        data_coeff#()
        
        'Choose csv file to use, multiple columns fine
        csvFilepath = modText.csvFind
        
        'If nothing chosen, close program
        If Len(csvFilepath) = 0 Then
            Exit Sub
        End If
        
        'Pull data from csv file into array, no assumption of columnar formatting
        csvMatrix = modText.csvParse(csvFilepath)
        
        'Separate x and y arrays from larger csvMatrix, if needed
        dataX = modMatrix.matVec(csvMatrix, 1)
        dataY = modMatrix.matVec(csvMatrix, 3)
        
        'Combine separate x and y arrays for use in optPolyFit which takes 2D (x,y) array
        data_2D = modMatrix.matJoin(dataX, dataY)
        
        'Generate n-order polynomial, columnar formatting assumed (c1 = x, c2 = y)
        data_Fit = modOptimization.optPolyFit(data_2D, 1)
        
        'Write the various results to written filepaths
        csvStatus = modText.csvWrite(csvMatrix, "raw.csv")
        csvStatus = modText.csvWrite(data_2D, "xy.csv")
        csvStatus = modText.csvWrite(data_Fit, "order1.csv")

End Sub

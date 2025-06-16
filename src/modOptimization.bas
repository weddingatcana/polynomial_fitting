Attribute VB_Name = "modOptimization"
Option Explicit

Public Function optPolyCoeff(ByRef A#(), _
                             ByVal polyOrder&) As Double()

    Dim rawX#(), _
        rawY#(), _
        Vm#(), _
        i_coeff#(), _
        f_coeff#()

        
        rawX = modMatrix.matVec(A, 1)
        rawY = modMatrix.matVec(A, 2)
        Vm = modMath.mathVandermonde(rawX, polyOrder)
        
        i_coeff = modMatrix.matPin(Vm)
        f_coeff = modMatrix.matMul(i_coeff, rawY)
        
        optPolyCoeff = f_coeff

End Function

Public Function optPolyFit(ByRef A#(), _
                           ByVal polyOrder&) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        rawMaxRowCoeff&, _
        coeff#(), _
        i&, k&, _
        sum#, _
        C#()
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        
        coeff = optPolyCoeff(A, polyOrder)
        rawMaxRowCoeff = UBound(coeff, 1)
        
        ReDim C(1 To rawMaxRowA, 1 To rawMaxColA)
        
        For i = 1 To rawMaxRowA
        
            sum = 0
            For k = 1 To rawMaxRowCoeff
            
                sum = sum + (coeff(k, 1) * (A(i, 1) ^ (k - 1)))
                    
            Next k
            
            C(i, 1) = A(i, 1)
            C(i, 2) = sum
            
        Next i
        
        optPolyFit = C

End Function

Public Function optPolyFit_seperate_coeff(ByRef A#(), _
                                          ByRef coeff#()) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        rawMaxRowCoeff&, _
        i&, k&, _
        sum#, _
        C#()
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        rawMaxRowCoeff = UBound(coeff, 1)
        
        ReDim C(1 To rawMaxRowA, 1 To (rawMaxColA + 1))
        
        For i = 1 To rawMaxRowA
        
            sum = 0
            For k = 1 To rawMaxRowCoeff
            
                sum = sum + (coeff(k, 1) * (A(i, 1) ^ (k - 1)))
                    
            Next k
            
            C(i, 1) = A(i, 1)
            C(i, 2) = sum
            
        Next i
        
        optPolyFit_seperate_coeff = C

End Function

Public Function optSSR#(ByRef A#(), _
                        ByRef B#())
                        
    Dim rawMaxRowA&, _
        rawMaxRowB&, _
        sum#, _
        i&
        
        rawMaxRowA = UBound(A, 1)
        rawMaxRowB = UBound(B, 1)
               
        If rawMaxRowA <> rawMaxRowB Then
            Exit Function
        End If
        
        sum = 0
        For i = 1 To rawMaxRowA
        
            sum = sum + (A(i, 1) - B(i, 1)) ^ 2
        
        Next i
        
        optSSR = sum

End Function

Public Function optSST#(ByRef A#(), _
                        ByVal avg#)

    Dim rawMaxRowA&, _
        sum#, _
        i&
        
        rawMaxRowA = UBound(A, 1)
        
        sum = 0
        For i = 1 To rawMaxRowA
        
            sum = sum + (A(i, 1) - avg) ^ 2
        
        Next i
        
        optSST = sum
        
End Function

Public Function optR2#(ByVal SSR#, _
                       ByVal SST#)
                       
    optR2 = 1 - (SSR / SST)
                       
End Function

Public Function optAvg#(ByRef A#())

    Dim rawMaxRowA&, _
        avg#, _
        sum#, _
        i&
        
        rawMaxRowA = UBound(A, 1)
        
        sum = 0
        For i = 1 To rawMaxRowA
            sum = sum + A(i, 1)
        Next i
        
        avg = sum / rawMaxRowA
        optAvg = avg

End Function

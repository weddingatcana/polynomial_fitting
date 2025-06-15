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

Public Function optSavGol(ByRef A#(), _
                 Optional ByVal window& = 11, _
                 Optional ByVal polyOrder& = 2) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        moving_mid&, _
        i&, j&, _
        k&, p&, _
        length&, _
        buffer#(), _
        Poly#(), _
        mid&, _
        C#()

        If window Mod 2 <> 0 And _
           window >= polyOrder + 1 Then
        Else
            Exit Function
        End If

        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        length = rawMaxRowA - (window - 1)
        mid = window \ 2
        moving_mid = mid
        
        If rawMaxColA <> 2 Then
            Exit Function
        End If
        
        ReDim C(1 To length, 1 To rawMaxColA)
        ReDim buffer(1 To window, 1 To rawMaxColA)
        ReDim Poly(1 To window, 1 To rawMaxColA)
        
        For i = 1 To length
            
            j = i
            k = 1
            Do
            
                If k > window Then
                    Exit Do
                End If
        
                buffer(k, 1) = A(j, 1)
                buffer(k, 2) = A(j, 2)
                
                j = j + 1
                k = k + 1
        
            Loop
            
            Poly = modOptimization.optPolyFit(buffer, polyOrder)
            
            C(i, 1) = A(moving_mid, 1)
            C(i, 2) = Poly(mid, 2)
            moving_mid = moving_mid + 1
        
        Next i

        optSavGol = C
        
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

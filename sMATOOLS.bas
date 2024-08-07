Attribute VB_Name = "sMATOOLS"
Option Explicit
Option Compare Text
' Option Base 1
' In general, range array base start with 1 and vba array with 0 base
' Excel/VBA reads array left to right and then top to bottom. To work with range array and vba array
' together, all array will be converted to base 1

' sMATTOOLS are a collection of matrix operation and compatible with MATLAB
' Author: Yeol C. Seong
' Date: 2001/08/01
' Contact: monkeyquant@gmail.com
' REVISED: 2013/10/04
'       - Added new comments to clarify the works and removed unnecessary comments
'       - optimized codes
'        : 2023/12/20
'       - Added Inverse using Application.MINVERSE
'       - Changed names to avoid any conflicts
'       - Reformatted samples in worksheets


Public Function elemOp(ByVal leftArr As Variant, ByVal rightArr As Variant, ByVal cKey As Variant) As Variant

    Dim op As String
    
    If TypeName(cKey) = "Range" Then
        op = LCase(cKey.Value)
        Set cKey = Nothing
    Else
        op = LCase(cKey)
        cKey = Empty
    End If
    
    Select Case True
        Case (op = "add" Or op = "plus" Or op = "+")
            elemOp = elemPlus(leftArr, rightArr)
        Case (op = "subtract" Or op = "minus" Or op = "-")
            elemOp = elemMinus(leftArr, rightArr)
        Case (op = "multiply" Or op = "product" Or op = "*")
            elemOp = elemMultiply(leftArr, rightArr)
        Case (op = "Divide" Or op = "/")
            elemOp = elemDivide(leftArr, rightArr)
        Case Else
            MsgBox "no such operation defined"
            elemOp = "#VALUE!"
    End Select

End Function



Public Function elemDivide(ByVal leftArr As Variant, ByVal rightArr As Variant) As Variant
' elemDIV: Calculate Matrix Division element by element.
'
' For Example, A={1,2; 3,4}, and B = {6,7;9,8}, Then B ./A = {6,3.5; 3,2}.
' If either number of column or rows are matching with those of array, then such row
' or column vector will be applied to entire rows or columns
'
' %%Author: Yeol C. Seong
' %%Date: 2001/08/01
' %%Contact: yseong@uchicago.edu or yeol.seong@bmo.com
' %%REVISED:
'

    Dim NoRowA As Variant, NoColA As Variant
    Dim NoRowB As Variant, NoColB As Variant
    Dim MAT_A As Variant, MAT_B As Variant
    Dim i As Integer, j As Integer
    Dim cAns() As Variant
    
    ' to be used as a Worksheet Function and VBA Array together
    ' all arrays from Rangestarting with base 1, not 0
    If TypeName(leftArr) = "Range" Then
        MAT_A = leftArr.Value
        Set leftArr = Nothing
        NoRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
        NoColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(leftArr) = 0 Then
            MAT_A = leftArr
            leftArr = Empty
            
            NoRowA = UBound(MAT_A, 1) - LBound(MAT_A, 1) + 1       ' Measuring The length of the rows of a Matrix
            NoColA = UBound(MAT_A, 2) - LBound(MAT_A, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_A(1 To NoRowA, 1 To NoColA)
        Else
            MAT_A = leftArr
            leftArr = Empty
            
            NoRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
            NoColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If
    
    ' to be used as a Worksheet Function and VBA Array together
    ' all arrays from Rangestarting with base 1, not 0
    If TypeName(rightArr) = "Range" Then
        MAT_B = rightArr.Value
        Set rightArr = Nothing
        NoRowB = UBound(MAT_B, 1)       ' Measuring The length of the rows of a Matrix
        NoColB = UBound(MAT_B, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(rightArr) = 0 Then
            MAT_B = rightArr
            rightArr = Empty
            
            NoRowB = UBound(MAT_B, 1) - LBound(MAT_B, 1) + 1       ' Measuring The length of the rows of a Matrix
            NoColB = UBound(MAT_B, 2) - LBound(MAT_B, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_A(1 To NoRowB, 1 To NoColB)
        Else
            MAT_B = rightArr
            rightArr = Empty
            
            NoRowB = UBound(MAT_B, 1)       ' Measuring The length of the rows of a Matrix
            NoColB = UBound(MAT_B, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If

    ' leftArray and rightArray both are matrix
    
    ' lieftArray is column or row vector, and vice versa
    
    ' one is row vector and the other is column vector
    
    ' else case
    
    If NoRowA = NoRowB Then
        If NoColA = NoColB Then
    
            ReDim Preserve MAT_A(1 To NoRowA, 1 To NoColA)
            ReDim Preserve MAT_B(1 To NoRowB, 1 To NoColB)
            ReDim cAns(1 To NoRowA, 1 To NoColA) As Variant
    
            For i = 1 To NoRowA
                For j = 1 To NoColA
                    If MAT_B(i, j) <> 0 Then
                        cAns(i, j) = MAT_A(i, j) / MAT_B(i, j)
                    Else
                        cAns(i, j) = "#DIV/0!"
                    End If
                Next j
            Next i
            elemDivide = cAns
        Else
            MsgBox "The size of matrix is not matched"
            elemDivide = "Erro 13"
        End If
    Else
        MsgBox "The size of matrix is not matched"
        elemDivide = "Error 13"
    
    End If

End Function




Public Function elemMultiply(ByVal leftArr As Variant, ByVal rightArr As Variant) As Variant
'
' elemPROD: Calculate Matrix Multiplication of element by element.
'
'   For Example, A={1,2; 3,4}, and B = {6,7;8,9}, Then Ans = {6,14; 24,36}.
'
'
' Author: Yeol C. Seong
' Date: 2001/08/01
' Contact: yseong@uchicago.edu or yeol.seong@bmo.com
' REVISED:
'
    
    Dim NoRowA As Variant, NoColA As Variant    ' size of left matrix
    Dim NoRowB As Variant, NoColB As Variant    ' size of right matrix
    Dim MAT_A As Variant, MAT_B As Variant      ' all arrrays including range into vba array
    Dim i As Integer, j As Integer, k As Integer    ' dummy variables
    Dim cAns() As Variant
    Dim tAns As Double                          ' column and row calculation
    
    ' to be used as a Worksheet Function and VBA Array together
    ' all arrays from Rangestarting with base 1, not 0
    If TypeName(leftArr) = "Range" Then
        MAT_A = leftArr.Value                ' When we use Array as Input.
        Set leftArr = Nothing                ' Free up memory
        NoRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
        NoColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(leftArr) = 0 Then
            MAT_A = leftArr
            leftArr = Empty
            NoRowA = UBound(MAT_A, 1) - LBound(MAT_A, 1) + 1       ' Measuring The length of the rows of a Matrix
            NoColA = UBound(MAT_A, 2) - LBound(MAT_A, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_A(1 To NoRowA, 1 To NoColA)
        Else
            MAT_A = leftArr
            leftArr = Empty
            NoRowA = UBound(MAT_A, 1)   ' Measuring The length of the rows of a Matrix
            NoColA = UBound(MAT_A, 2)   ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If
    
    ' to be used as a Worksheet Function and VBA Array together
    ' all arrays from Rangestarting with base 1, not 0
    If TypeName(rightArr) = "Range" Then
        MAT_B = rightArr.Value
        Set rightArr = Nothing
        NoRowB = UBound(MAT_B, 1)       ' Measuring The length of the rows of a Matrix
        NoColB = UBound(MAT_B, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(rightArr) = 0 Then
            MAT_B = rightArr
            rightArr = Empty
            NoRowB = UBound(MAT_B, 1) - LBound(MAT_B, 1) + 1       ' Measuring The length of the rows of a Matrix
            NoColB = UBound(MAT_B, 2) - LBound(MAT_B, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_B(1 To NoRowB, 1 To NoColB)
        Else
            MAT_B = rightArr
            rightArr = Empty
            NoRowB = UBound(MAT_B, 1)   ' Measuring The length of the rows of a Matrix
            NoColB = UBound(MAT_B, 2)   ' Measuring The length of the size of columns of a Matrix
        End If
    End If

    If NoRowA = NoRowB Then
        If NoColA = NoColB Then
            ReDim Preserve MAT_A(1 To NoRowA, 1 To NoColA)
            ReDim Preserve MAT_B(1 To NoRowB, 1 To NoColB)
            ReDim cAns(1 To NoRowA, 1 To NoColA) As Variant
                
            For i = 1 To NoRowA
                For j = 1 To NoColB
                        cAns(i, j) = MAT_A(i, j) * MAT_B(i, j)
                Next j
            Next i
        
            elemMultiply = cAns
        Else
            MsgBox "Error 13: The size of matrix is not matched"
            elemMultiply = "Error 13"
        End If
    Else
        MsgBox "Error 13: The size of matrix is not matched"
        elemMultiply = "Error 13"
    End If
    
End Function




Public Function elemPlus(ByVal leftArr As Variant, ByVal rightArr As Variant) As Variant
'
' elemAdd: Calculate Matrix Multiplication of element by element.
'
'   For Example, A={1,2; 3,4}, and B = {6,7;8,9}, Then Ans = {6,14; 24,36}.
'
'
' Author: Yeol C. Seong
' Date: 2001/08/01
' Contact: yseong@uchicago.edu or yeol.seong@bmo.com
' REVISED:
'
    
    Dim NoRowA As Variant, NoColA As Variant    ' size of left matrix
    Dim NoRowB As Variant, NoColB As Variant    ' size of right matrix
    Dim MAT_A As Variant, MAT_B As Variant      ' all arrrays including range into vba array
    Dim i As Integer, j As Integer, k As Integer    ' dummy variables
    Dim cAns() As Variant
    Dim tAns As Double                          ' column and row calculation
    
    ' to be used as a Worksheet Function and VBA Array together
    ' all arrays from Rangestarting with base 1, not 0
    If TypeName(leftArr) = "Range" Then
        MAT_A = leftArr.Value                ' When we use Array as Input.
        Set leftArr = Nothing                ' Free up memory
        NoRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
        NoColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(leftArr) = 0 Then
            MAT_A = leftArr
            leftArr = Empty
            NoRowA = UBound(MAT_A, 1) - LBound(MAT_A, 1) + 1       ' Measuring The length of the rows of a Matrix
            NoColA = UBound(MAT_A, 2) - LBound(MAT_A, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_A(1 To NoRowA, 1 To NoColA)
        Else
            MAT_A = leftArr
            leftArr = Empty
            NoRowA = UBound(MAT_A, 1)   ' Measuring The length of the rows of a Matrix
            NoColA = UBound(MAT_A, 2)   ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If
    
    ' to be used as a Worksheet Function and VBA Array together
    ' all arrays from Rangestarting with base 1, not 0
    If TypeName(rightArr) = "Range" Then
        MAT_B = rightArr.Value
        Set rightArr = Nothing
        NoRowB = UBound(MAT_B, 1)       ' Measuring The length of the rows of a Matrix
        NoColB = UBound(MAT_B, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(rightArr) = 0 Then
            MAT_B = rightArr
            rightArr = Empty
            NoRowB = UBound(MAT_B, 1) - LBound(MAT_B, 1) + 1       ' Measuring The length of the rows of a Matrix
            NoColB = UBound(MAT_B, 2) - LBound(MAT_B, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_B(1 To NoRowB, 1 To NoColB)
        Else
            MAT_B = rightArr
            rightArr = Empty
            NoRowB = UBound(MAT_B, 1)   ' Measuring The length of the rows of a Matrix
            NoColB = UBound(MAT_B, 2)   ' Measuring The length of the size of columns of a Matrix
        End If
    End If

    If NoRowA = NoRowB Then
        If NoColA = NoColB Then
            ReDim Preserve MAT_A(1 To NoRowA, 1 To NoColA)
            ReDim Preserve MAT_B(1 To NoRowB, 1 To NoColB)
            ReDim cAns(1 To NoRowA, 1 To NoColA) As Variant
                
            For i = 1 To NoRowA
                For j = 1 To NoColB
                        cAns(i, j) = MAT_A(i, j) + MAT_B(i, j)
                Next j
            Next i
        
            elemPlus = cAns
        Else
            MsgBox "Error 13: The size of matrix is not matched"
            elemPlus = "Error 13"
        End If
    Else
        MsgBox "Error 13: The size of matrix is not matched"
        elemPlus = "Error 13"
    End If
    
End Function




Public Function elemMinus(ByVal leftArr As Variant, ByVal rightArr As Variant) As Variant
'
' elemPROD: Calculate Matrix Multiplication of element by element.
'
'   For Example, A={1,2; 3,4}, and B = {6,7;8,9}, Then Ans = {6,14; 24,36}.
'
'
' Author: Yeol C. Seong
' Date: 2001/08/01
' Contact: yseong@uchicago.edu or yeol.seong@bmo.com
' REVISED:
'
    
    Dim NoRowA As Variant, NoColA As Variant    ' size of left matrix
    Dim NoRowB As Variant, NoColB As Variant    ' size of right matrix
    Dim MAT_A As Variant, MAT_B As Variant      ' all arrrays including range into vba array
    Dim i As Integer, j As Integer, k As Integer    ' dummy variables
    Dim cAns() As Variant
    Dim tAns As Double                          ' column and row calculation
    
    ' to be used as a Worksheet Function and VBA Array together
    ' all arrays from Rangestarting with base 1, not 0
    If TypeName(leftArr) = "Range" Then
        MAT_A = leftArr.Value                ' When we use Array as Input.
        Set leftArr = Nothing                ' Free up memory
        NoRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
        NoColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(leftArr) = 0 Then
            MAT_A = leftArr
            leftArr = Empty
            NoRowA = UBound(MAT_A, 1) - LBound(MAT_A, 1) + 1       ' Measuring The length of the rows of a Matrix
            NoColA = UBound(MAT_A, 2) - LBound(MAT_A, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_A(1 To NoRowA, 1 To NoColA)
        Else
            MAT_A = leftArr
            leftArr = Empty
            NoRowA = UBound(MAT_A, 1)   ' Measuring The length of the rows of a Matrix
            NoColA = UBound(MAT_A, 2)   ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If
    
    ' to be used as a Worksheet Function and VBA Array together
    ' all arrays from Rangestarting with base 1, not 0
    If TypeName(rightArr) = "Range" Then
        MAT_B = rightArr.Value
        Set rightArr = Nothing
        NoRowB = UBound(MAT_B, 1)       ' Measuring The length of the rows of a Matrix
        NoColB = UBound(MAT_B, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(rightArr) = 0 Then
            MAT_B = rightArr
            rightArr = Empty
            NoRowB = UBound(MAT_B, 1) - LBound(MAT_B, 1) + 1       ' Measuring The length of the rows of a Matrix
            NoColB = UBound(MAT_B, 2) - LBound(MAT_B, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_B(1 To NoRowB, 1 To NoColB)
        Else
            MAT_B = rightArr
            rightArr = Empty
            NoRowB = UBound(MAT_B, 1)   ' Measuring The length of the rows of a Matrix
            NoColB = UBound(MAT_B, 2)   ' Measuring The length of the size of columns of a Matrix
        End If
    End If

    If NoRowA = NoRowB Then
        If NoColA = NoColB Then
            ReDim Preserve MAT_A(1 To NoRowA, 1 To NoColA)
            ReDim Preserve MAT_B(1 To NoRowB, 1 To NoColB)
            ReDim cAns(1 To NoRowA, 1 To NoColA) As Variant
                
            For i = 1 To NoRowA
                For j = 1 To NoColB
                        cAns(i, j) = MAT_A(i, j) - MAT_B(i, j)
                Next j
            Next i
        
            elemMinus = cAns
        Else
            MsgBox "Error 13: The size of matrix is not matched"
            elemMinus = "Error 13"
        End If
    Else
        MsgBox "Error 13: The size of matrix is not matched"
        elemMinus = "Error 13"
    End If
    
End Function




Function matSize(ByVal targetArr As Variant, Optional ByVal cKey As Variant = -1) As Variant
' %% matSize: results the size of array, mimic MATLAB size
' %% Outputs:
'       one of array
' %% Input:
'       cArr                Array - vector or matrix
'       cKey:
'           -1 or 'all'     for array of all dimmension
'           1 or 'row'      for row
'           2 or 'column'   for column
'           ...
'
' %%Author: Yeol C. Seong
' %%Date: 2001/08/01
' %%Contact: yseong@uchicago.edu or yeol.seong@bmo.com
' %%REVISED:
'

    Dim tArr As Variant
    
    If TypeName(targetArr) = "Range" Then
        tArr = targetArr.Value
        Set targetArr = Nothing
    Else
        tArr = targetArr
        targetArr = Empty
    End If
        
    On Error Resume Next
    Select Case IIf(IsNumeric(cKey), cKey, LCase(Left(cKey, 1)))
        Case 1, "r"
            sSize = UBound(tArr, 1) + IIf(LBound(tArr, 1) = 0, 1, 0)
        Case 2, "c"
             sSize = UBound(tArr, 2) + IIf(LBound(tArr, 1) = 0, 1, 0)
        Case -1, "a"
            Dim tI As Long
            tI = 1
            Dim tX() As Variant
            
            Do
                ReDim Preserve tX(1 To 1, 1 To tI)
                tX(1, tI) = UBound(tArr, tI) + IIf(LBound(tArr, tI) = 0, 1, 0)
                tI = tI + 1
            Loop Until IsError(UBound(tArr, tI))

        '' Another Do While Loop Approach
        ''    Dim i As Integer
        ''
        ''    GetArrDim = -99
        ''    If Not IsArray(inArr) Then Exit Function
        ''
        ''    On Error Resume Next
        ''        Err.Clear
        ''        Do While IsNumeric(UBound(inArr, i + 1))
        ''            If Err.Number = 9 Then Exit Do
        ''            i = i + 1
        ''        Loop
        ''    GetArrDim = i
        ''    On Error GoTo 0
 
            matSize = tX
        Err.Clear
    End Select
    
''    matSize = UBound(arr, whichDimension) + IIf(LBound(arr) = 0, 1, 0)
    On Error GoTo 0
End Function



Public Function matResize(ByVal targetArr As Variant, ByVal nRows As Long, Optional ByVal nCols As Long = 1) As Variant
' Resize the given Array by given row and column size in the column based order.
' Total number of elements should be the same as row x column size.
' Default: Make matrix a column vector, mimic MATLAB reshape.
' All resize logic is from top to bottom and then move to next column
'
'   Inputs
'
' Author:     Yeol C. Seong
' Date:       2001/08/01
' Contact:    yseong@uchicago.edu
' REVISED:
'

' %% Making this function available for both Worksheet and VBA
    Dim tArr As Variant
    
    If TypeName(targetArr) = "Range" Then
        tArr = targetArr.Value
        Set targetArr = Nothing
    Else
        tArr = targetArr
        targetArr = Empty
    End If

    ' Number of rows, columns and elements in the source matrix
    Dim tRows As Long, tCols As Long, tE As Long
    Dim i As Long, j As Long
    
    tRows = UBound(tArr, 1) + IIf(LBound(tArr, 1) = 0, 1, 0)
    tCols = UBound(tArr, 2) + IIf(LBound(tArr, 2) = 0, 1, 0)
    
    tE = tRows * tCols
    Dim tN As Long
    tN = nRows * nCols
    If (tE < tN) Or Not IsArray(tArr) Then
        MsgBox ("Out of Scope")
        Exit Function
    End If
         
    Dim tempArray As Variant
    tempArray = sArray2Vector(tArr, True)
    
    tArr = Empty
    
    Dim tArray() As Variant
    ReDim tArray(1 To nRows, 1 To nCols)
    Dim tCount As Long
    
    For j = 1 To nCols
        For i = 1 To nRows
           tCount = tCount + 1
            tArray(i, j) = tempArray(tCount, 1)
        Next i
    Next j

    matResize = tArray
End Function



Function repeatSequence(ByVal cArr As Variant, ByVal nRepeat As Long, Optional ByVal cDirection As Variant = "row") As Variant
' %%                                        Repeat and resize the given Array.

' %%
'   Inputs
'
' %%Author: Yeol C. Seong
' %%Date: 2001/08/01
' %%Contact: yseong@uchicago.edu or yeol.seong@bmo.com
' %%REVISED:
'

    Dim nRow As Variant, nCol As Variant
    Dim MAT_A As Variant, tDirection As Variant
    Dim i As Long, j As Long, k As Long
    Dim cAns As Variant
    Dim tRepeat As Long
    Dim tAddtional As Long
    Dim nLength As Long

    
    If TypeName(cArr) = "Range" Then
        MAT_A = cArr.Value
        Set cArr = Nothing
        If Not IsArray(MAT_A) Then MsgBox ("This is not array"): Exit Function

        nRow = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
        nCol = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(cArr) = 0 Then
            MAT_A = cArr
            Erase cArr
            If Not IsArray(MAT_A) Then MsgBox ("This is not array"): Exit Function
            
            nRow = UBound(MAT_A, 1) - LBound(MAT_A, 1) + 1       ' Measuring The length of the rows of a Matrix
            nCol = UBound(MAT_A, 2) - LBound(MAT_A, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_A(1 To nRow, 1 To nCol)
        Else
            MAT_A = cArr
            Erase cArr
            If Not IsArray(MAT_A) Then MsgBox ("This is not array"): Exit Function
            
            nRow = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
            nCol = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If
    
    If TypeName(cDirection) = "Range" Then
        tDirection = cDirection.Value
        Set cDirection = Nothing
    Else
        tDirection = cDirection
    End If
    If Not IsNumeric(tDirection) Then
        If LCase(Left(tDirection, 1)) = "c" Then                  ' To the right
            tDirection = 2
        Else                                                ' To down
            tDirection = 1
        End If
    End If
'
    Select Case tDirection
        Case 2                                              'Column
            If nRepeat < nCol Then
                sRepeatSequence = sResize(MAT_A, nRow, nRepeat)
            Else
                
                If (nRepeat Mod nCol) = 0 Then
                    tRepeat = nRepeat \ nCol - 1              ' Repeat number, "\" is integer division, the same as INT(a/b) or FLOOR(A/b)
                Else
                    tRepeat = nRepeat \ nCol
                End If
                
''                tLength = nCol * tRepeat
''                tLength = nCol * (1 + tRepeat)
''                tAddtional = nRepeat - tRepeat * nCol
''
                '' cAns = MAT_A
                '' ReDim Preserve cAns(1 To nRow, 1 To tLength)
''                ReDim cAns(1 To nRow, 1 To tLengh)
''
''                For k = 1 To tRepeat
''                    For j = 1 To nCol
''                        For i = 1 To nRow
''                            cAns(i, j + nCol) = MAT_B(i, j)
''                        Next i
''                    Next j
''                Next k
                
                cAns = MAT_A

                For k = 1 To tRepeat
                    cAns = sAppendArray(cAns, MAT_A, tDirection)
                Next k
                
                ReDim Preserve cAns(1 To nRow, 1 To nRepeat)
                repeatSequence = cAns
            End If
            
        Case Else                       ' Row Direction

''                cAns = Application.WorksheetFunction.Transpose(MAT_A)
''
''                ReDim Preserve cAns(1 To nColA, 1 To nRowA + nRowB)
''
''                For i = 1 To nRowB
''                    For j = 1 To nColA
''                        cAns(j, i + nRowA) = MAT_B(i, j)
''                    Next j
''                Next i
''
''                sRepeatSequence = Application.WorksheetFunction.Transpose(cAns)

            If nRepeat < nRow Then
                repeatSequence = sResize(MAT_A, nRepeat, nCol)
            Else
                
                If (nRepeat Mod nRow) = 0 Then
                    tRepeat = nRepeat \ nRow - 1              ' Repeat number, "\" is integer division, the same as INT(a/b) or FLOOR(A/b)
                Else
                    tRepeat = nRepeat \ nRow
                End If
                
                cAns = MAT_A
                For k = 1 To tRepeat
                    cAns = sAppendArray(cAns, MAT_A, tDirection)
                Next k
                
                cAns = Application.WorksheetFunction.Transpose(cAns)
                ReDim Preserve cAns(1 To nCol, 1 To nRepeat)
''                ReDim Preserve MAT_A(1 To nRepeat, 1 To nCol)
                repeatSequence = Application.WorksheetFunction.Transpose(cAns)
            End If
               
    End Select

End Function




Public Function createArray(Optional ByVal numRows As Variant, Optional ByVal numColumns As Variant, Optional zeroBased As Boolean = True) As Variant
' createArray: creates m x n matrix as a placeholder with various either zero or 1 based
' this will be used mostly in VBA, not in Worksheet.

    Dim tempArray As Variant
    
    If IsMissing(numRows) Then
        If IsMissing(numColumns) Then
            createArray = 0
            Exit Function
        Else
            If zeroBased Then
                ReDim tempArray(numColumns - 1)
            Else
                ReDim tempArray(1 To numColumns)
            End If
        End If
    Else
         If IsMissing(numColumns) Then
            If zeroBased Then
                ReDim tempArray(numRows - 1, 0)
            Else
                ReDim tempArray(1 To numRows, 1 To 1)
            End If
        Else
            If zeroBased Then
                ReDim tempArray(numRows - 1, numColumns - 1)
            Else
                ReDim tempArray(1 To numRows, 1 To numColumns)
            End If
        End If
    End If
    
    
    createArray = tempArray
End Function


Function appendArray(ByVal cArr As Variant, ByVal cArrToAppend As Variant, Optional ByVal cDirection As Variant = "Row") As Variant
'' Append new same size Array to the existing array. If not the same, then adjusted to the first one
'
' %%Author: Yeol C. Seong
' %%Date: 2001/08/01
' %%Contact: yseong@uchicago.edu or yeol.seong@bmo.com
' %%REVISED:
'

    Dim nRowA As Variant, nColA As Variant
    Dim nRowB As Variant, nColB As Variant
    Dim MAT_A As Variant, MAT_B As Variant, tDirection As Variant
    Dim i As Long, j As Long
    Dim cAns() As Variant
    
    If TypeName(cArr) = "Range" Then
        MAT_A = cArr.Value
        Set cArr = Nothing
        If Not IsArray(MAT_A) Then MsgBox ("This is not array"): Exit Function

        nRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
        nColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(cArr) = 0 Then
            MAT_A = cArr
            Erase cArr
            If Not IsArray(MAT_A) Then MsgBox ("This is not array"): Exit Function
            
            nRowA = UBound(MAT_A, 1) - LBound(MAT_A, 1) + 1       ' Measuring The length of the rows of a Matrix
            nColA = UBound(MAT_A, 2) - LBound(MAT_A, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_A(1 To nRowA, 1 To nColA)
        Else
            MAT_A = cArr
            Erase cArr
            If Not IsArray(MAT_A) Then MsgBox ("This is not array"): Exit Function
            
            nRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
            nColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If
    
    If TypeName(cArrToAppend) = "Range" Then
        MAT_B = cArrToAppend.Value
        Set cArrToAppend = Nothing
        If Not IsArray(MAT_B) Then MsgBox ("This is not array"): Exit Function
        
        nRowB = UBound(MAT_B, 1)       ' Measuring The length of the rows of a Matrix
        nColB = UBound(MAT_B, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(cArrToAppend) = 0 Then
            MAT_B = cArrToAppend
            Erase cArrToAppend
            If Not IsArray(MAT_B) Then MsgBox ("This is not array"): Exit Function
            
            nRowB = UBound(MAT_B, 1) - LBound(MAT_B, 1) + 1       ' Measuring The length of the rows of a Matrix
            nColB = UBound(MAT_B, 2) - LBound(MAT_B, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_A(1 To nRowB, 1 To nColB)
        Else
            MAT_B = cArrToAppend
            Erase cArrToAppend
            If Not IsArray(MAT_B) Then MsgBox ("This is not array"): Exit Function
            
            nRowB = UBound(MAT_B, 1)       ' Measuring The length of the rows of a Matrix
            nColB = UBound(MAT_B, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If

    If TypeName(cDirection) = "Range" Then
        tDirection = cDirection.Value
        Set cDirection = Nothing
    Else
        tDirection = cDirection
    End If
    If Not IsNumeric(tDirection) Then
        If LCase(Left(tDirection, 1)) = "c" Then                  ' To the right
            tDirection = 2
        Else                                                ' To down
            tDirection = 1
        End If
    End If
'
    Select Case tDirection
        Case 2                                              'Column
            If nRowA <> nRowB Then
                
                cAns = MAT_A
                ReDim Preserve cAns(1 To nRowA, 1 To nColA + nColB)
                
                For j = 1 To nColB
                    For i = 1 To nRowA
                        If i > nRowB Then
                            cAns(i, j + nColA) = 0
                        Else
                            cAns(i, j + nColA) = MAT_B(i, j)
                        End If
                    Next i
                Next j
                
                appendArray = cAns

            Else
                cAns = MAT_A
                ReDim Preserve cAns(1 To nRowA, 1 To nColA + nColB)
                
                For j = 1 To nColB
                    For i = 1 To nRowA
                        cAns(i, j + nColA) = MAT_B(i, j)
                    Next i
                Next j
                
                appendArray = cAns

            End If
            
        Case Else
            If nColA <> nColB Then
               ' MsgBox ("both dimensions should be the same. Resize the array to append")
                
                cAns = Application.WorksheetFunction.Transpose(MAT_A)
                
                ReDim Preserve cAns(1 To nColA, 1 To nRowA + nRowB)
                
                For i = 1 To nRowB
                    For j = 1 To nColA
                        If j > nColB Then
                            cAns(j, i + nRowA) = 0
                        Else
                            cAns(j, i + nRowA) = MAT_B(i, j)
                        End If
                    Next j
                Next i
                
                appendArray = Application.WorksheetFunction.Transpose(cAns)
                
            Else
                cAns = Application.WorksheetFunction.Transpose(MAT_A)
                
                ReDim Preserve cAns(1 To nColA, 1 To nRowA + nRowB)
                
                For i = 1 To nRowB
                    For j = 1 To nColA
                        cAns(j, i + nRowA) = MAT_B(i, j)
                    Next j
                Next i
                
                appendArray = Application.WorksheetFunction.Transpose(cAns)
                
            End If
    End Select


End Function



Function prependArray(ByVal cArr As Variant, ByVal cArrToPrepend As Variant, Optional ByVal cDirection As Variant = "Row") As Variant
'' prepend new same size Array to the existing array. If not the same, then adjusted to the first one
        
        
' %%Author: Yeol C. Seong
' %%Date: 2001/08/01
' %%Contact: yseong@uchicago.edu or yeol.seong@bmo.com
' %%REVISED:
'

    Dim nRowA As Variant, nColA As Variant
    Dim nRowB As Variant, nColB As Variant
    Dim MAT_A As Variant, MAT_B As Variant, tDirection As Variant
    Dim i As Long, j As Long
    Dim cAns() As Variant
    
    If TypeName(cArr) = "Range" Then
        MAT_A = cArr.Value
        Set cArr = Nothing
        If Not IsArray(MAT_A) Then MsgBox ("This is not array"): Exit Function

        nRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
        nColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        MAT_A = cArr
        cArr = Empty
    
        If LBound(MAT_A) = 0 Then
            If Not IsArray(MAT_A) Then MsgBox ("This is not array"): Exit Function
            
            nRowA = UBound(MAT_A, 1) - LBound(MAT_A, 1) + 1       ' Measuring The length of the rows of a Matrix
            nColA = UBound(MAT_A, 2) - LBound(MAT_A, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_A(1 To nRowA, 1 To nColA)
        Else
            If Not IsArray(MAT_A) Then MsgBox ("This is not array"): Exit Function
            
            nRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
            nColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If
    
    If TypeName(cArrToPrepend) = "Range" Then
        MAT_B = cArrToPrepend.Value
        Set cArrToPrepend = Nothing
        If Not IsArray(MAT_B) Then MsgBox ("This is not array"): Exit Function
        
        nRowB = UBound(MAT_B, 1)       ' Measuring The length of the rows of a Matrix
        nColB = UBound(MAT_B, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        MAT_B = cArrToPrepend
        cArrToPrepend = Empty
        If LBound(MAT_B) = 0 Then
            If Not IsArray(MAT_B) Then MsgBox ("This is not array"): Exit Function
            
            nRowB = UBound(MAT_B, 1) - LBound(MAT_B, 1) + 1       ' Measuring The length of the rows of a Matrix
            nColB = UBound(MAT_B, 2) - LBound(MAT_B, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_A(1 To nRowB, 1 To nColB)
        Else
            If Not IsArray(MAT_B) Then MsgBox ("This is not array"): Exit Function
            
            nRowB = UBound(MAT_B, 1)       ' Measuring The length of the rows of a Matrix
            nColB = UBound(MAT_B, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If

    If TypeName(cDirection) = "Range" Then
        tDirection = cDirection.Value
        Set cDirection = Nothing
    Else
        tDirection = cDirection
        cDirection = Empty
    End If
    
    If Not IsNumeric(tDirection) Then
        If LCase(Left(tDirection, 1)) = "c" Then                  ' To the right
            tDirection = 2
        Else                                                ' To down
            tDirection = 1
        End If
    End If
'
    Select Case tDirection
        Case 2                                              'Column
            If nRowA <> nRowB Then
                
                cAns = MAT_B
                ReDim Preserve cAns(1 To nRowB, 1 To nColA + nColB)
                
                For j = 1 To nColA
                    For i = 1 To nRowB
                        If i > nRowA Then
                            cAns(i, j + nColB) = 0
                        Else
                            cAns(i, j + nColB) = MAT_A(i, j)
                        End If
                    Next i
                Next j
                
                prependArray = cAns

            Else
                cAns = MAT_B
                ReDim Preserve cAns(1 To nRowB, 1 To nColA + nColB)
                
                For j = 1 To nColA
                    For i = 1 To nRowB
                        cAns(i, j + nColB) = MAT_A(i, j)
                    Next i
                Next j
                
                prependArray = cAns

            End If
            
        Case Else
            If nColA <> nColB Then
                
                cAns = Application.WorksheetFunction.Transpose(MAT_B)
                
                ReDim Preserve cAns(1 To nColB, 1 To nRowA + nRowB)
                
                For i = 1 To nRowA
                    For j = 1 To nColB
                        If j > nColA Then
                            cAns(j, i + nRowB) = 0
                        Else
                            cAns(j, i + nRowB) = MAT_A(i, j)
                        End If
                    Next j
                Next i
                
                prependArray = Application.WorksheetFunction.Transpose(cAns)
                
            Else
                cAns = Application.WorksheetFunction.Transpose(MAT_B)
                
                ReDim Preserve cAns(1 To nColB, 1 To nRowA + nRowB)
                
                For i = 1 To nRowA
                    For j = 1 To nColB
                        cAns(j, i + nRowB) = MAT_A(i, j)
                    Next j
                Next i
                
                prependArray = Application.WorksheetFunction.Transpose(cAns)
                
            End If
    End Select


End Function



Function Array_Prepend(arr As Variant, ParamArray valuesToPrepend() As Variant) As Variant
' By JP Technologies

    Dim tempArray As Variant
    Dim i As Long   ' make temp array the same size as
    ' the current array + number of new elements
    ReDim tempArray(LBound(arr) To (UBound(arr) + UBound(valuesToPrepend) + 1))
    ' put new elements at the beginning of the new array
    For i = LBound(valuesToPrepend) To UBound(valuesToPrepend)
        tempArray(i) = valuesToPrepend(i)
    Next i   ' put existing elements after the new ones
    For i = UBound(valuesToPrepend) + 1 To UBound(tempArray)
        tempArray(i) = arr(i - (UBound(valuesToPrepend) + 1))
    Next i
    Array_Prepend = tempArray
End Function
'
''
Public Function sRowSize(ByVal cA As Variant) As Variant
''    If IsObject(sArray) Then
''        sRowSize = sArray.Rows.Count
''    ElseIf IsArray(sArray) Then
''        sRowSize = UBound(sArray, 1) - LBound(sArray, 1) + 1
''    Else
''        Debug.Print "Input is neither range nor array."
''
''        Exit Function
''
''    End If

    Dim nRowA As Long, nColA As Long
    Dim tA As Variant
    
    If TypeName(cA) = "Range" Then
        tA = cA.Value
        Set cA = Nothing
        nRowA = UBound(tA, 1)       ' Measuring The length of the rows of a Matrix
        nColA = UBound(tA, 2)       ' Measuring The length of the size of columns of a Matrix
    Else
        If LBound(cA) = 0 Then
            tA = cA
            Erase cA
            
            nRowA = UBound(tA, 1) - LBound(tA, 1) + 1       ' Measuring The length of the rows of a Matrix
            nColA = UBound(tA, 2) - LBound(tA, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
        Else
            tA = cA
            Erase cA
            
            nRowA = UBound(tA, 1)       ' Measuring The length of the rows of a Matrix
            nColA = UBound(tA, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
    End If
    
    sRowSize = nRowA
    
End Function


'
''
Public Function sColumnSize(ByVal cA As Variant) As Variant
''    If IsObject(sArray) Then
''        sColumnSize = sArray.Columns.Count
''    ElseIf IsArray(sArray) Then
''        sColumnSize = UBound(sArray, 1) - LBound(sArray, 1) + 1
''    Else
''        Debug.Print "Input is neither range nor array."
''
''        Exit Function
''
''    End If

    Dim nRowA As Long, nColA As Long
    Dim tA As Variant
    
    If TypeName(cA) = "Range" Then
        tA = cA.Value
        Set cA = Nothing
        nRowA = UBound(tA, 1)       ' Measuring The length of the rows of a Matrix
        nColA = UBound(tA, 2)       ' Measuring The length of the size of columns of a Matrix
    Else
        If LBound(cA) = 0 Then
            tA = cA
            Erase cA
            
            nRowA = UBound(tA, 1) - LBound(tA, 1) + 1       ' Measuring The length of the rows of a Matrix
            nColA = UBound(tA, 2) - LBound(tA, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
        Else
            tA = cA
            Erase cA
            
            nRowA = UBound(tA, 1)       ' Measuring The length of the rows of a Matrix
            nColA = UBound(tA, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
    End If
    
    sColumnSize = nColA


End Function

'
''
' %% Count total number of entries under certain criterion
Public Function sDarwinCount(ByVal sArray As Variant, Optional ByVal sIndex As Integer = 1, _
                    Optional ByVal sKey As Boolean) As Variant
' %% sDarwinCount counts certain group of entries from sArray
'                 returns two column vectors
'
'       sIndex      Set to the first column
'       sKey        True for Row Count, and False for Column Count
        
    Dim i, j, k, l  As Integer
    Dim sCount, tValue As Integer
    Dim sRows, sColumns As Integer
    
    If IsError(UBound(sArray, 1)) Then
        If IsError(UBound(sArray, 2)) Then
            Exit Function
        Else
            sRows = 1
            sColumns = UBound(sArray, 2) - LBound(sArray, 2) + 1
        End If
    Else
        If IsError(UBound(sArray, 2)) Then
            sRows = UBound(sArray, 1) - LBound(sArray, 1) + 1
            sColumns = 1
        Else
            sRows = UBound(sArray, 1) - LBound(sArray, 1) + 1
            sColumns = UBound(sArray, 2) - LBound(sArray, 2) + 1
        End If
    End If
        
    Dim tmpValues() As Variant
    ReDim tmpValues(1 To sRows, 1 To 2)
    
    If sKey Then
        sArray = sSortByVal(sArray, sIndex)
        For i = 1 To sRows - 1
            k = k + 1
            sCount = 0
            For j = i + 1 To sRows
                If sArray(i, sIndex) = sArray(j, sIndex) Then
                    sCount = sCount + 1
                    tValue = sCount + 1
                End If
            Next j
                tmpValues(k, 1) = sArray(i, sIndex)
                If sCount > 0 Then
                    tmpValues(k, 2) = tValue
                Else
                    tmpValues(k, 2) = 1
                End If
                i = i + sCount
        Next i
    Else
            sArray = WorksheetFunction.Transpose(sSortByVal(WorksheetFunction.Transpose(sArray), sIndex))
         For i = 1 To sColumns - 1
            k = k + 1
            sCount = 0
            For j = i + 1 To sColumns
                If sArray(sIndex, i) = sArray(sIndex, j) Then
                    sCount = sCount + 1
                    tValue = sCount + 1
                End If
            Next j
                tmpValues(1, k) = sArray(sIndex, i)
                If sCount > 0 Then
                    tmpValues(2, k) = tValue
                Else
                    tmpValues(2, k) = 1
                End If
                i = i + sCount
        Next i
   
    End If
    sDarwinCount = sRemoveEmptyCells(tmpValues, sIndex, sKey)
End Function
' End sDarwinCount



'
''
' %% Count total number of entries under certain criterion
Public Function sDarwinCountShort(ByVal sArray As Variant, Optional sIndex As Integer = 1) As Variant
' %% sDarwinCount counts certain group of entries from sArray
'
'
        
    Dim i, j, k, l  As Integer
    Dim sCount, tValue As Integer
    Dim sRows, sColumns As Integer
    
    sArray = sRemoveEmptyCells(sArray)
    If sMATOOLS.IsArrayEmpty(sArray) Then
        Exit Function
    Else
        If IsError(UBound(sArray, 2)) Then
            sRows = UBound(sArray, 1) - LBound(sArray, 1) + 1
            sColumns = 1
        Else
            sRows = UBound(sArray, 1) - LBound(sArray, 1) + 1
            sColumns = UBound(sArray, 2) - LBound(sArray, 2) + 1
        End If
    End If
    
    Dim tmpValues() As Variant
   
    ReDim tmpValues(1 To sRows, 1 To 2)
            
    sArray = sSortByVal(sArray, sIndex)

    For i = 1 To sRows - 1
        
        k = k + 1
        sCount = 0
        For j = i + 1 To sRows
            If sArray(i, sIndex) = sArray(j, sIndex) Then
                sCount = sCount + 1
                tValue = sCount + 1
             End If
               
        Next j
            tmpValues(k, 1) = sArray(i, sIndex)
            
            If sCount > 0 Then
                tmpValues(k, 2) = tValue
            Else
                tmpValues(k, 2) = 1
            End If

            i = i + sCount
   
    Next i
    
    'sDarwinCountShort = sRemoveEmptyCells(tmpValues)
    Dim tmpResult As Variant
    ReDim tmpResult(1 To k, 1 To 2)
    For l = 1 To k
        For j = 1 To 2
            tmpResult(l, j) = tmpValues(l, j)
        Next j
    Next l
    sDarwinCountShort = tmpResult
    
End Function
' End sDarwinCountShort



'
''
'''
Public Function sIndexMatch4Row(ByVal sLOOKUP As Variant, ByVal sMatrix As Variant, ByVal sColRef As Integer, Optional ByVal sRefArray As Variant)
' s_IndexMatch will replace Vlookup Function
'
'       Input:
'                   sLookup             the value to look for in sMatrix and sRefArray
'                   SMatrix             the reference pool of data
'                                       the first column will be used for a Index Lookup Reference if sRefArr is missing
'                   sColRef             the reference column which may contain data looking for
'                   sRefArr             the reference column, Optional


Dim tmpAns As Variant           ' %% result
Dim sRow1, sCol1, sRow2, sCol2 As Integer
sRow1 = UBound(sMatrix, 1): sCol1 = UBound(sMatrix, 2)



If IsEmpty(sRefArr) Then
    tmpAns = Applicaiton.Index(Application.Index(sLOOKUP, , 1), Application.Match(sLOOKUP, sMatrix, sColCount, 0), 0)
Else

    sRow2 = UBound(sRefArr, 1): sCol2 = UBound(sRefArr, 2)
    
    If sRow1 <> sRow2 Then Exit Function

    tmpAns = Application.Index(sRefArr, Application.Match(sLOOKUP, sMatrix, sColCount, 0), 0)

End If

    sIndexMatch4Row = tmpAns

End Function


'
''
'''
Public Function sSortByVal(ByVal sArray As Variant, Optional ByVal ColIndex As Integer) As Variant
' %% s_SortByVal
    Dim i, j, k As Integer
    Dim TempValue() As Variant
    Dim sRows, sColumns As Integer
    
    If sMATOOLS.IsArrayEmpty(sArray) Then
        Exit Function
    Else
        If IsError(UBound(sArray, 2)) Then
            sRows = UBound(sArray, 1) - LBound(sArray, 1) + 1
            sColumns = 1
        Else
            sRows = UBound(sArray, 1) - LBound(sArray, 1) + 1
            sColumns = UBound(sArray, 2) - LBound(sArray, 2) + 1
        End If
    End If


' %% If no ColIndex was specified, then this function use the first column
If ColIndex = Null Then
    ColIndex = 1
End If

ReDim TempValue(1 To UBound(sArray, 2))

For i = 1 To UBound(sArray, 1) - 1
    For j = i + 1 To UBound(sArray, 1)
        If sArray(i, ColIndex) > sArray(j, ColIndex) Then
            For k = 1 To UBound(sArray, 2)
                TempValue(k) = sArray(j, k)
                sArray(j, k) = sArray(i, k)
                sArray(i, k) = TempValue(k)
            Next k
        End If
    Next j
Next i

    sSortByVal = sArray
    
End Function


'
''
'''
Public Function sArray2Vector(ByVal cArr As Variant, Optional ByVal cKey As Boolean = True) As Variant
' %% sARray2Vector makes retangular matrix (Two Arrays of Calendar) to a single vector
' Read Column after column or row after row
'
'   Inputs:
'           sArray
'           sKey         :  1 or True for Column vector
'                           0 or False for Row Vector

' %%Author: Yeol C. Seong
' %%Date: 2001/08/01
' %%Contact: yseong@uchicago.edu or yeol.seong@bmo.com
' %%REVISED:
'

    Dim tArr As Variant
    
    If TypeName(cArr) = "Range" Then
        tArr = cArr.Value
    Else
        tArr = cArr
    End If
        
    On Error Resume Next

    Dim i, j, k, RowCount, ColumnCount, NewSize As Integer
    Dim tmpValues() As Variant
    
        If IsArrayEmpty(tArr) Then
                Exit Function
        Else
            If IsError(UBound(tArr, 2)) Then                        ' No column size, so column vector
                RowCount = UBound(tArr, 1) - LBound(tArr, 1) + 1
                ColumnCount = 1
            ElseIf IsError(UBound(tArr, 1)) Then
                RowCount = 1
                ColumnCount = UBound(tArr, 2) - LBound(tArr, 2) + 1
            Else
                RowCount = UBound(tArr, 1) - LBound(tArr, 1) + 1
                ColumnCount = UBound(tArr, 2) - LBound(tArr, 2) + 1
            End If
        End If

        NewSize = RowCount * ColumnCount
    
        If cKey Then
    
            ReDim tmpValues(1 To NewSize, 1 To 1)
        
            For j = 1 To ColumnCount
                For i = 1 To RowCount
      
                    k = k + 1
                    tmpValues(k, 1) = tArr(i, j)
            
                Next i
            Next j
    
        Else
    
            ReDim tmpValues(1 To 1, 1 To NewSize)
        
            For j = 1 To ColumnCount
                For i = 1 To RowCount
                    k = k + 1
                    tmpValues(1, k) = tArr(i, j)
                Next i
            Next j
    
        End If
    
        sArray2Vector = tmpValues

End Function

'
''
' %%
Public Function sRemoveEmptyCells(ByVal cArr As Variant, Optional ByVal cIndex As Variant = 1, _
        Optional ByVal cKey As Variant = "column", Optional ByVal cValue As Variant = "string") As Variant
' %% sREMOVEEMPTYCELLS              removes any empty or zeros in Array/Range
'
' Inputs:
'   sIndex                          Determine which reference row or column to be reviewed
'   sKey = "row" or 1               Operation direction: default: removes all empty cells in Rows ideentified by Index
'          "column" or 2
'   sValue = "string" or 0          for empty cell and False for zero
'           "numeric" or 1
'           "any" or -1
'
' Only take in consideration of a

    Dim tArr As Variant
    Dim sIndex As Variant
    Dim sKey As Variant
    Dim sValue As Variant
    Dim myArray() As Variant
    Dim i, j, l, k, nRow, nCol As Integer

    '' Convert Range into VBA Array with Base = 1
    If TypeName(cArr) = "Range" Then
        tArr = cArr.Value
        Set cArr = Nothing
        
        If sMATOOLS.IsArrayEmpty(tArr) Then
            Exit Function
        Else
            If IsError(UBound(tArr, 2)) Then
                nRow = UBound(tArr, 1) - LBound(tArr, 1) + 1
                nCol = 1
            Else
                ' Range is always referred to Base 1
                nRow = UBound(tArr, 1)       ' Measuring The length of the rows of a Matrix
                nCol = UBound(tArr, 2)       ' Measuring The length of the size of columns of a Matrix
            End If
        End If
    '' VBA Array Case, in general, VBA Array start its indexing from 0. Make it to Base = 1
    Else
        tArr = cArr
        Erase cArr
        If sMATOOLS.IsArrayEmpty(tArr) Then
            Exit Function
        Else
            If LBound(tArr) = 0 Then
                If IsError(UBound(tArr, 2)) Then
                    nRow = UBound(tArr, 1) - LBound(tArr, 1) + 1
                    nCol = 1
                    
                    ReDim Preserve tArr(1 To nRow, 1 To nCol)
                ElseIf IsError(UBound(tArr, 1)) Then
                    nCol = UBound(tArr, 1) - LBound(tArr, 1) + 1
                    nRow = 1
                    
                    ReDim Preserve tArr(1 To nRow, 1 To nCol)
                Else
                    nRow = UBound(tArr, 1) - LBound(tArr, 1) + 1
                    nCol = UBound(tArr, 2) - LBound(tArr, 2) + 1
                    
                    ReDim Preserve tArr(1 To nRow, 1 To nCol)
                End If
            Else
                nRow = UBound(tArr, 1)       ' Measuring The length of the rows of a Matrix
                nCol = UBound(tArr, 2)       ' Measuring The length of the size of columns of a Matrix
                
                ReDim Preserve tArr(1 To nRow, 1 To nCol)
            End If
            
        End If
    End If
    
    If TypeName(cIndex) = "Range" Then
        sIndex = cIndex.Value
        Set cIndex = Nothing
    Else
        sIndex = cIndex
    End If
    
    If TypeName(cKey) = "Range" Then
        sKey = cKey.Value
        Set cKey = Nothing
    Else
        sKey = cKey
    End If
    
    If Not IsNumeric(sKey) Then
        If LCase(Left(sKey, 1)) = "c" Then
            sKey = 2
        Else
            sKey = 1
        End If
    End If

    If TypeName(cValue) = "Range" Then
        sValue = cValue.Value
        Set cValue = Nothing
    Else
        sValue = cValue
    End If
    
    If Not IsNumeric(sValue) Then
        If LCase(Left(sValue, 1)) = "a" Then
            sValue = -1
        ElseIf LCase(Left(sValue, 1)) = "n" Then
            sValue = 1
        Else
            sValue = 0                          ' Default for string
        End If
    End If

    
        
''    myArray = sSortByVal(myArray, sIndex)
        
    Select Case sKey
        Case 2                                  ' Column direction
            '' Due to ReDim only applicable the last dimension
            ReDim myArray(1 To nRow, 1 To nCol) As Variant

            Select Case sValue
                Case 1                          ' Numeric
                    For k = 1 To nCol
                        On Error Resume Next
                        If IsNumeric(tArr(sIndex, k)) And tArr(sIndex, k) = 0 Then
                             l = nCol - 1
                             ReDim Preserve myArray(1 To nRow, 1 To l)
                         Else
                             j = j + 1
                             For i = 1 To nRow
                                 myArray(i, j) = tArr(i, k)
                             Next i
                         End If
                         On Error GoTo 0
                    Next k
            
                Case -1                         ' All - zero or empty string
                    For k = 1 To nCol
                        On Error Resume Next
                        If IsEmpty(tArr(sIndex, k)) Or (IsNumeric(tArr(sIndex, k)) And tArr(sIndex, k) = 0) Then
                             l = nCol - 1
                             ReDim Preserve myArray(1 To nRow, 1 To l)
                         Else
                             j = j + 1
                             For i = 1 To nRow
                                 myArray(i, j) = tArr(i, k)
                             Next i
                         End If
                         On Error GoTo 0
                    Next k
               
                Case Else                       ' Empty String
                    For k = 1 To nCol
                        On Error Resume Next
                        If IsEmpty(tArr(sIndex, k)) Then
                            l = nCol - 1
                            ReDim Preserve myArray(1 To nRow, 1 To l)
                        Else
                            j = j + 1
                            For i = 1 To nRow
                                myArray(i, j) = tArr(i, k)
                            Next i
                        End If
                        On Error GoTo 0
                    Next k
            End Select

        Case Else                               ' Row Direction
            Select Case sValue
                '' Due to ReDim only applicable the last dimension
                ReDim myArray(1 To nCol, 1 To nRow) As Variant
                Case 1                          ' numeric zeros
                    For i = 1 To nRow
                        On Error Resume Next
                        If IsNumeric(tArr(i, sIndex)) And tArr(i, sIndex) = 0 Then
                            l = nRow - 1
                            ReDim Preserve myArray(1 To nCol, 1 To l)
                        Else
                            j = j + 1
                            For k = 1 To nCol
                                myArray(k, j) = tArr(i, k)
                            Next k
                        End If
                        On Error GoTo 0
                    Next i
                    myArray = Application.Transpose(myArray)
                Case -1                         ' both zeros or empty
                    For i = 1 To nRow
                        On Error Resume Next
                        If IsEmpty(tArr(i, sIndex)) Or (IsNumeric(tArr(i, sIndex)) And tArr(i, sIndex) = 0) Then
                            l = nRow - 1
                            ReDim Preserve myArray(1 To nCol, 1 To l)
                        Else
                            j = j + 1
                            For k = 1 To nCol
                                myArray(k, j) = tArr(i, k)
                            Next k
                        End If
                        On Error GoTo 0
                    Next i
                    myArray = Application.Transpose(myArray)
              
                Case Else                       ' empty string
                    For i = 1 To nRow
                        On Error Resume Next
                        If IsEmpty(tArr(i, sIndex)) Then
                            l = nRow - 1
                            ReDim Preserve myArray(1 To nCol, 1 To l)
                        Else
                            j = j + 1
                            For k = 1 To nCol
                                myArray(k, j) = tArr(i, k)
                            Next k
                        End If
                        On Error GoTo 0
                    Next i
                    myArray = Application.Transpose(myArray)

            End Select
    End Select
        
    sRemoveEmptyCells = myArray
    

End Function
' %% End sRomoveEmptyCells




'' %% Dumplicate Date Remover
Public Function sDuplicateRemoverOld(ByVal sArray As Variant, Optional ByVal sIndex As Integer = 1) As Variant
' %% sDuplicateRemover will remove any duplicate entries in vector
'   Inputs
'               sArray     A column Vector


    Dim i1, i2, i3, j1, j2, j3, k1, k2, k3, iRowCount, jRowCount, sRows, sColumns As Integer
        sRows = UBound(sArray, 1) - LBound(sArray, 1) + 1
        sColumns = UBound(sArray, 2) - LBound(sArray, 2) + 1

    Dim tmpValues() As Variant
    Dim tmpValue() As Variant

    ReDim tmpValues(1 To sRows, 1 To sColumns)

    For k1 = 1 To sColumns
        tmpValues(1, k1) = sArray(1, k1)
    Next k1
    
    iRowCount = 1
    Do
        i1 = i + 1
        For j1 = i1 + 1 To sRows
            If tmpValues(i1, sIndex) <> sArray(j1, sIndex) Then
                    iRowCount = iRowCount + 1
                For k2 = 1 To sColumns
                    tmpValues(iRowCount, k2) = sArray(j1, k2)
                Next k2
            Else
                jRowCount = jRowCount + 1
            End If
        Next j1
        
        Dim tmpArray() As Variant
        ReDim tmpArray(1 To iRowCount, 1 To sColumns)
        For i2 = 1 To iRowCount
            For j2 = 1 To sColumns
                tmpArray(i2, j2) = tmpValues(i2, j2)
            Next j2
        Next i2
        'ReDim Preserve TmpValues(1 To iRowCount, 1 To sColumns)
        'sDuplicateRemover = sDuplicateRemover(tmpArray, sIndex)
        
    Loop Until (i1 <= sRows Or jRowCount < 1)
    
    sDuplicateRemoverOld = tmpArray
End Function
' %% End sDuplicateRemoverOld




'' %% Dumplicate Date Remover
Public Function sDuplicateRemover(ByVal sArray As Variant, Optional ByVal sIndex As Integer = 1, _
    Optional ByVal sKey As Boolean = True) As Variant
' %% sDuplicateRemover will remove any duplicate entries in vector
'   Inputs
'               sArray     A matrix
'               sIndex      Column or Row
'               sKey        Yes for Row Removal and No for Column Removal

    Dim i, j, k, sRows, sColumns, iCounter As Integer
    Dim RowCount, ColumnCount As Integer
    If IsError(UBound(sArray, 1)) Then
        If IsError(UBound(sArray, 2)) Then
            Exit Function
        Else
            RowCount = 1
            ColumnCount = UBound(sArray, 2) - LBound(sArray, 2) + 1
        End If
    Else
        If IsError(UBound(sArray, 2)) Then
            RowCount = UBound(sArray, 1) - LBound(sArray, 1) + 1
            ColumnCount = 1
        Else
            RowCount = UBound(sArray, 1) - LBound(sArray, 1) + 1
            ColumnCount = UBound(sArray, 2) - LBound(sArray, 2) + 1
        End If
    End If
   
    sArray = sSortByVal(sArray, sIndex)
    
    Dim tmpValues() As Variant

    ReDim tmpValues(1 To RowCount, 1 To ColumnCount)
    
    sRows = RowCount
    sColumns = ColumnCount
    
    If sKey Then                ' Remove Duplicates in Rows
        If RowCount <> 1 Then
            For l = 1 To ColumnCount
                tmpValues(1, l) = sArray(1, l)
            Next l
            For i = 1 To sRows - 1
                For j = i + 1 To sRows
                    If (sArray(i, sIndex) = sArray(j, sIndex)) Then
                        For k = j To sRows - 1
                            For l = 1 To ColumnCount
                                sArray(k, l) = sArray(k + 1, l)
                            Next l
                        Next k
                                    
                        For l = 1 To ColumnCount
                            sArray(sRows, l) = " "
                        Next l
                                    
                        iCounter = iCounter + 1

                        sRows = sRows - 1
                
                        j = j - 1
                    End If
                Next j
        'Debug.Print sVector
        
            Next i
        Else            ' If input row is a single entry
            Exit Function
        End If
    Else                ' Column Removes
        If Column <> 1 Then     ' If input columns are more than one
            For l = 1 To RowCount
                tmpValues(l, 1) = sArray(l, 1)
            Next l
            For i = 1 To sColumns - 1
                For j = i + 1 To sColumns
                    If (sArray(sIndex, i) = sArray(sIndex, j)) Then
                        For k = j To sColumns - 1
                            For l = 1 To RowCount
                                sArray(l, k) = sArray(l, k + 1)
                            Next l
                        Next k
                                    
                        For l = 1 To RowCount
                            sArray(l, sColumns) = " "
                        Next l
                                    
                        iCounter = iCounter + 1

                        sColumns = sColumns - 1
                
                        j = j - 1
                    End If
                Next j
        'Debug.Print sVector
        
            Next i
        Else
            Exit Function
        End If
    End If
     
    Dim sResult() As Variant
    If sKey Then
        ReDim sResult(1 To RowCount - iCounter, 1 To ColumnCount)
        For i = 1 To RowCount - iCounter
            For j = 1 To ColumnCount
                sResult(i, j) = sArray(i, j)
            Next j
        Next i
    Else
        ReDim sResult(1 To RowCount, 1 To ColumnCount - iCounter)
        For j = 1 To ColumnCount - iCounter
            For i = 1 To RowCount
                sResult(i, j) = sArray(i, j)
            Next i
        Next j
    End If
    
    sDuplicateRemover = sResult

End Function
' %% End sDuplicateRemover




'' %% Dumplicate Date Remover
Public Function sDuplicateRemover2(ByVal sArray As Variant, Optional ByVal sIndex As Integer = 1) As Variant
' %% sDuplicateRemover will remove any duplicate entries in vector
'   Inputs
'               sArray     A column Vector


    Dim i, i1, i2, j, j1, j2, k, l, iRowCount, jRowCount, sRows, sColumns As Integer
        sRows = UBound(sArray, 1) - LBound(sArray, 1) + 1
        sColumns = UBound(sArray, 2) - LBound(sArray, 2) + 1

    Dim tmpValues() As Variant
    Dim TmpRowValues() As Variant

    ReDim tmpValues(1 To sRows)
    
    ' %% Allocate the first entry
    For j1 = 1 To sColumns
        tmpValues(1, j1) = sArray(1, j1)
    Next j1
    
    For i = 1 To sRows - 1
        For j = i + 1 To sRows
            If sArray(i, sIndex) = sArray(j, sIndex) Then
                
                For k = j To sRows - 1
                    For l = 1 To sColumns
                        sArray(k, l) = sArray(k + 1, l)
                    Next l
                Next k
                
                    For l = 1 To sColumns
                        sArray(sRows, l) = " "
                    Next l
                iRowCount = iRowCount + 1
                sRows = sRows - 1
                j = j - 1
                'ReDim Preserve sArray(1 To sRows, 1 To sColumns)
              
            End If
        Next j
        
        
    Next i
    
    sDuplicateRemover2 = sRemoveEmptyCells(sVector, sIndex)

End Function
' %% End sDuplicateRemover2




' %% extracting certain row(s) or column(s) from matrix
Public Function sExtractColumnPartitionedMatrix(ByVal sArray As Variant, ByVal sIndex As Variant) As Variant
' Extract Certain Identified Columns from Matrix

    Dim i, j, RowCount, ColumnCount As Integer
    
    If IsError(UBound(sArray, 1)) Then
        If IsError(UBound(sArray, 2)) Then
            Exit Function
        Else
            RowCount = 1
            ColumnCount = UBound(sArray, 2) - LBound(sArray, 2) + 1
        End If
    Else
        If IsError(UBound(sArray, 2)) Then
            RowCount = UBound(sArray, 1) - LBound(sArray, 1) + 1
            ColumnCount = 1
        Else
            RowCount = UBound(sArray, 1) - LBound(sArray, 1) + 1
            ColumnCount = UBound(sArray, 2) - LBound(sArray, 2) + 1
        End If
    End If
    

    If IsArray(sIndex) Then
        IndexCount = UBound(sIndex) - LBound(sIndex) + 1
    Else
        IndexCount = 1
    End If
    Dim tmpValues() As Variant
    Dim T As Variant
    Dim tMax As Long
    tMax = WorksheetFunction.Max(sIndex)
    
    If ColumnCount < tMax Then
        Debug.Print "The Column, " & j & ", you want to extract out is not in the scope. We end this process ..."
        Exit Function
    End If
    
    ReDim tmpValues(1 To RowCount, 1 To IndexCount)
    
    If IndexCount > 1 Then
    For j = 1 To IndexCount
        If IsEmpty(sIndex(j)) Then
            Debug.Print "The Row, " & j & ", you want to extract out is not in the scope. We end this process ..."
            Exit Function
        Else
            For i = 1 To RowCount
                tmpValues(i, j) = sArray(i, sIndex(j))
            Next i
        End If
    Next j
    
    Else
        If IsEmpty(sIndex) Then
            Exit Function
        Else
            For i = 1 To RowCount
                tmpValues(i, 1) = sArray(1, sIndex)
            Next i
        End If
    End If
    sExtractColumnPartitionedMatrix = tmpValues
    
End Function
' %% End sExtractColumnPartitionedMatrix


'
''
' %% extracting certain row(s) or column(s) from matrix
Public Function sExtractRowPartitionedMatrix(ByVal sArray As Variant, ByVal sIndex As Variant) As Variant
' Extract Certain Identified Rows from Matrix
'           sIndex is a vector

    Dim i, j, k, kk, RowCount, ColumnCount As Integer
    If IsError(UBound(sArray, 1)) Then
        If IsError(UBound(sArray, 2)) Then
            Exit Function
        Else
            RowCount = 1
            ColumnCount = UBound(sArray, 2) - LBound(sArray, 2) + 1
        End If
    Else
        If IsError(UBound(sArray, 2)) Then
            RowCount = UBound(sArray, 1) - LBound(sArray, 1) + 1
            ColumnCount = 1
        Else
            RowCount = UBound(sArray, 1) - LBound(sArray, 1) + 1
            ColumnCount = UBound(sArray, 2) - LBound(sArray, 2) + 1
        End If
    End If
    
    
    If IsArray(sIndex) Then
        IndexCount = UBound(sIndex) - LBound(sIndex) + 1
    Else
        IndexCount = 1
    End If
    Dim tmpValues() As Variant
    Dim tMax As Long
    tMax = WorksheetFunction.Max(sIndex)
    
    If RowCount < tMax Then
        Debug.Print "The Column, " & j & ", you want to extract out is not in the scope. We end this process ..."
        Exit Function
    End If
    
    ReDim tmpValues(1 To IndexCount, 1 To ColumnCount)
    'ReDim Preserve sIndex(1 To IndexCount)
    
    
    If IndexCount > 1 Then
        For i = 1 To IndexCount
            If IsEmpty(sIndex(i)) Then
                Debug.Print "The Row, " & j & ", you want to extract out is not in the scope. We end this process ..."
                Exit Function
            Else
                For j = 1 To ColumnCount
                    tmpValues(i, j) = sArray(sIndex(i), j)
                Next j
            End If
        Next i
    Else
        If IsEmpty(sIndex) Then
            Exit Function
        Else
            For j = 1 To ColumnCount
                tmpValues(1, j) = sArray(sIndex, j)
            Next j
        End If
    End If
    sExtractRowPartitionedMatrix = tmpValues
    

End Function
' %% End sExtractRowPartitionedMatrix



'
''
' %% Union in VBA Array
Public Function sUnion(ByVal sKey As Boolean, ParamArray sArray() As Variant) As Variant
' sUNION combines two or more matrices in row or column by sKey

' %%Author: Yeol C. Seong
' %%Date: 2001/08/01
' %%Contact: yseong@uchicago.edu or yeol.seong@bmo.com
' %%REVISED:
'

''    Dim tArr As Variant
''
''    If TypeName(cArr) = "Range" Then
''        tArr = cArr.Value
''    Else
''        tArr = cArr
''    End If

    Dim i, j, k, l As Integer
    Dim rMaxLoop(), cMaxLoop() As Variant
    Dim tmpSize(), tmpValues() As Variant
    Dim RowsInput, ColumnsInput As Integer
    RowsInput = UBound(sArray, 1) - LBound(sArray, 1) + 1
    ReDim tmpSize(1 To RowsInput, 1 To 2)
    
    For i = LBound(sArray) To UBound(sArray)
        Dim tmpValue() As Variant
        tmpValue = sArray(i)
        tmpSize(i + 1, 1) = UBound(tmpValue, 1) - LBound(tmpValue, 1) + 1
        tmpSize(i + 1, 2) = UBound(tmpValue, 2) - LBound(tmpValue, 2) + 1
    Next i
        
    'rMaxLoop = sMathTools.sRowColumnSum(tmpSize, True)
    cMaxLoop = sMathTools.sRowColumnSum(tmpSize, False)
    
    
    ReDim tmpValues(1 To cMaxLoop(1, 1), 1 To cMaxLoop(1, 2))
    
    For i = LBound(sArray) To UBound(sArray)
            Dim tmpArraysInArray() As Variant
            tmpArraysInArray = sArray(i)
            Dim RowSize, ColumnSize As Variant
            RowSize = UBound(tmpArraysInArray, 1) - LBound(tmpArraysInArray, 1) + 1
            ColumnSize = UBound(tmpArraysInArray, 2) - LBound(tmpArraysInArray, 2) + 1
        
        If sKey Then    ' %% Row expansion
                For j = 1 To RowSize
                    For k = l + 1 To ColumnSize
                        tmpValues(j, k) = tmpArraysInArray(j, k)
                    Next k
                Next j
                    l = l + ColumnSize
     
        Else                ' %% Column expansion
                For j = l + 1 To RowSize
                    For k = 1 To ColumnSize
                        tmpValues(j, k) = tmpArraysInArray(j, k)
                    Next k
                Next j
                        l = l + RowSize
        End If
    Next i
    
    If sKey Then
        ReDim Preserve tmpValues(1 To l, 1 To cMaxLoop(1, 2))
    Else
        ReDim Preserve tmpValues(1 To rMaxLoop(1, 1), 1 To l)
    End If
    
    sUnion = tmpValues
    
End Function
' End sUnion

'
''
' %% Union in VBA Array
Public Function sUnion2(ParamArray sArray() As Variant) As Variant
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Modified C Pearson's Union2 by Y Seong on 2/15/2015
' A Union operation that accepts parameters that are Nothing.
' This for Array in VBA, not Range
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer
    Dim RR As Range
    Dim SS() As Variant
    
 
        For i = LBound(sArray) To UBound(sArray)
            If IsObject(sArray(i)) Then
                If Not sArray(i) Is Nothing Then
                    If TypeOf sArray(i) Is Excel.Range Then
                        If Not RR Is Nothing Then
                            Set sUnion2 = Application.Union(RR, sArray(i))
                        Else
                            Set sUnion2 = sArray(i)
                        End If
                    End If
                End If
            End If
        Next i
    
End Function
' End sUnion2
     
     
 Function ProperUnion(ParamArray Ranges() As Variant) As Range
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ProperUnion
' This provides Union functionality without duplicating
' cells when ranges overlap. Requires the Union2 function.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Source: http://cpearson.com/excel/BetterUnion.aspx

    Dim ResR As Range
    Dim N As Long
    Dim R As Range
    
    If Not Ranges(LBound(Ranges)) Is Nothing Then
        Set ResR = Ranges(LBound(Ranges))
    End If
    For N = LBound(Ranges) + 1 To UBound(Ranges)
        If Not Ranges(N) Is Nothing Then
            For Each R In Ranges(N).Cells
                If Application.Intersect(ResR, R) Is Nothing Then
                    Set ResR = Union2(ResR, R)
                End If
            Next R
        End If
    Next N
    Set ProperUnion = ResR
End Function


Function Union2(ParamArray Ranges() As Variant) As Range
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Union2
' A Union operation that accepts parameters that are Nothing.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Source: http://cpearson.com/excel/BetterUnion.aspx

    Dim N As Long
    Dim RR As Range
    For N = LBound(Ranges) To UBound(Ranges)
        If IsObject(Ranges(N)) Then
            If Not Ranges(N) Is Nothing Then
                If TypeOf Ranges(N) Is Excel.Range Then
                    If Not RR Is Nothing Then
                        Set RR = Application.Union(RR, Ranges(N))
                    Else
                        Set RR = Ranges(N)
                    End If
                End If
            End If
        End If
    Next N
    Set Union2 = RR
End Function

Sub TestUnion2()
    Dim R1 As Range
    Dim R2 As Range
    Dim R3 As Range
    Set R3 = Range("J9:L14")
    Dim RR As Range
    Set R1 = Range("B2:C4")
    Set R2 = Range("E2:F4")
    'Set RR = Union2(R1, R2) ' success
    
    'Range("R1").Value = RR
    
    'Range("R10").Value = sRemoveEmptyCells(R3.Value, 1, 1, 1)
    
    If IsEmpty("") Then MsgBox ("Yes")
    
End Sub




'
''
' %% Union in VBA Array and Expand Rows
Public Function sTwoMatricesUnion(ByVal a As Variant, ByVal b As Variant, ByVal sKey As Boolean) As Variant
' sRowsUnion combines two matrices in row or column by sKey
'            combines rows by default

    Dim i, j, k, l As Integer
    
    Dim tmpMaxLoop As Long
    Dim tmpValues() As Variant
    Dim sParser As Variant

    Dim aRowsInput, aColumnsInput As Integer
    If sMATOOLS.IsArrayEmpty(a) Then
        If sMATOOLS.IsArrayEmpty(b) Then
            Exit Function
        Else
            If IsError(UBound(b, 2)) Then
                bRowsInput = UBound(b, 1) - LBound(b, 1) + 1
                bColumnsInput = 1
            Else
                bRowsInput = UBound(b, 1) - LBound(b, 1) + 1
                bColumnsInput = UBound(b, 2) - LBound(b, 2) + 1
            End If
        End If
           
    Else
        If sMATOOLS.IsArrayEmpty(b) Then
        
            If IsError(UBound(a, 2)) Then
                aRowsInput = UBound(a, 1) - LBound(a, 1) + 1
                aColumnsInput = 1
            Else
                aRowsInput = UBound(a, 1) - LBound(a, 1) + 1
                aColumnsInput = UBound(a, 2) - LBound(a, 2) + 1
            End If
        Else
            aRowsInput = UBound(a, 1) - LBound(a, 1) + 1
            aColumnsInput = UBound(a, 2) - LBound(a, 2) + 1
            bRowsInput = UBound(b, 1) - LBound(b, 1) + 1
            bColumnsInput = UBound(b, 2) - LBound(b, 2) + 1
        End If
       
    End If
        
   If Val(aRowsInput) > 1 Then
        If Val(bRowsInput) > 1 Then
            sParser = "AB"
        Else
            sParser = "AO"
        End If
    Else
        If Val(bRowsInput) > 1 Then
            sParser = "BO"
        Else
            sParser = "AB"
        End If
    End If
            
 
    Dim RowsInputMax, ColumnsInputMax As Integer
    
    Select Case sParser
    
        Case "AO"         ' A Only
            ReDim tmpValues(1 To aRowsInput, 1 To aColumnsInput)
            If sKey Then    ' %% Row expansion
                For i = 1 To aRowsInput
                    For j = 1 To aColumnsInput
                        If Val(a(i, j)) <> 0 Then
                            tmpValues(i, j) = a(i, j)
                        Else
                            tmpValues(i, j) = ""
                        End If
                    Next j
                Next i
                
            Else                ' %% Column expansion
                For j = 1 To aColumnsInput
                    For i = 1 To aRowsInput
                        If Val(a(i, j)) <> 0 Then
                            tmpValues(i, j) = a(i, j)
                        Else
                            tmpValues(i, j) = ""
                        End If
                    Next i
                Next j
            End If
        
        Case "BO"         ' B Only"
            ReDim tmpValues(1 To bRowsInput, 1 To bColumnsInput)
            
            If sKey Then    ' %% Row expansion
                For i = 1 To bRowsInput
                    For j = 1 To bColumnsInput
                        If Val(b(i, j)) <> 0 Then
                            tmpValues(i, j) = b(i, j)
                        Else
                            tmpValues(i, j) = ""
                        End If
                    Next j
                Next i
                
            Else                ' %% Column expansion
                For j = 1 To bColumnsInput
                    For i = 1 To bRowsInput
                        If Val(b(i, j)) <> 0 Then
                            tmpValues(i, j) = b(i, j)
                        Else
                            tmpValues(i, j) = ""
                        End If
                    Next i
                Next j
            End If
        
        Case "AB"         ' Both has non-empty rows
            If aRowsInput >= bRowsInput Then
                RowsInputMax = aRowsInput
            Else
                RowsInputMax = bRowsInput
            End If
    
            If aColumnsInput >= bColumnsInput Then
                ColumnsInputMax = aColumnsInput
            Else
                ColumnsInputMax = bColumnsInput
            End If
    
  
            If sKey Then
                tmpMaxLoop = aRowsInput + bRowsInput
                'ReDim Preserve tmpValues(1 To tmpMaxLoop, 1 To ColumnsInputMax)
                ReDim tmpValues(1 To tmpMaxLoop, 1 To ColumnsInputMax)

            Else
                tmpMaxLoop = aColumnsInput + bColumnsInput
                'ReDim Preserve tmpValues(1 To RowsInputMax, 1 To tmpMaxLoop)
                ReDim tmpValues(1 To RowsInputMax, 1 To tmpMaxLoop)
            End If
    
    
            If sKey Then    ' %% Row expansion
                For i = 1 To aRowsInput
                    For j = 1 To aColumnsInput
                        If Val(a(i, j)) <> 0 Then
                            tmpValues(i, j) = a(i, j)
                        Else
                            tmpValues(i, j) = ""
                        End If
                    Next j
                Next i
                
                For i = aRowsInput + 1 To tmpMaxLoop
                    k = k + 1
                    For j = 1 To bColumnsInput
                        If Val(b(k, j)) <> 0 Then
                            tmpValues(i, j) = b(k, j)
                        Else
                            tmpValues(i, j) = b(k, j)
                        End If
                    Next j
                Next i
            Else                ' %% Column expansion
                For j = 1 To aColumnsInput
                    For i = 1 To aRowsInput
                        If Val(a(i, j)) <> 0 Then
                            tmpValues(i, j) = a(i, j)
                        Else
                            tmpValues(i, j) = ""
                        End If
                    Next i
                Next j
                For j = aColumnsInput + 1 To tmpMaxLoop
                    k = k + 1
                    For i = 1 To bRowsInput
                        If Val(b(i, k)) <> 0 Then
                            tmpValues(i, j) = b(i, k)
                        Else
                            tmpValues(i, j) = ""
                        End If
                    Next i
                Next j
            End If
        
        Case Else
            Exit Function
    
    End Select
    
    
    sTwoMatricesUnion = tmpValues
    
End Function
' End sTwoMatricesUnion


'
''
Public Function sReplaceColumns(ByVal a As Variant, ByVal ref2A As Variant, ByVal b As Variant, Optional ByVal ref2B As Variant)
' This function swap culumns of A with those of B specified
' the number of columns to swap in A should be the same as those of B
' if the row size is different, truncated into the row number of A

    Dim i, j As Integer
    Dim aRows, bRows, aColumns, bColumns As Long
    Dim arRows, brRows, arColumns, brColumns As Long

    If IsArrayEmpty(a) Then
        If IsError(UBound(a, 2)) Then
            Exit Function
        Else
            aRows = 1
            aColumns = UBound(a, 2) - LBound(a, 2) + 1
        End If
    Else
        If IsError(UBound(a, 2)) Then
            aRows = UBound(a, 1) - LBound(a, 1) + 1
            aColumns = 1
        Else
            aRows = UBound(a, 1) - LBound(a, 1) + 1
            aColumns = UBound(a, 2) - LBound(a, 2) + 1
        End If
    End If
  If IsArrayEmpty(b) Then
        If IsError(UBound(b, 2)) Then
            Exit Function
        Else
            bRows = 1
            bColumns = UBound(b, 2) - LBound(b, 2) + 1
        End If
    Else
        If IsError(UBound(b, 2)) Then
            bRows = UBound(b, 1) - LBound(b, 1) + 1
            bColumns = 1
        Else
            bRows = UBound(b, 1) - LBound(b, 1) + 1
            bColumns = UBound(b, 2) - LBound(b, 2) + 1
        End If
    End If
    If IsArrayEmpty(ref2A) Then
        If IsError(UBound(ref2A, 2)) Then
            Exit Function
        Else
            arRows = 1
            arColumns = UBound(ref2A, 2) - LBound(ref2A, 2) + 1
        End If
    Else
        If IsError(UBound(ref2A, 2)) Then
            arRows = UBound(ref2A, 1) - LBound(ref2A, 1) + 1
            arColumns = 1
        Else
            arRows = UBound(ref2A, 1) - LBound(ref2A, 1) + 1
            arColumns = UBound(ref2A, 2) - LBound(ref2A, 2) + 1
        End If
    End If
    
    If IsArrayEmpty(ref2B) Then
        brRows = arRows
        brColumns = arColumns
        ref2B = ref2A
    Else
        If IsError(UBound(ref2B, 2)) Then
            brRows = UBound(ref2B, 1) - LBound(ref2B, 1) + 1
            brColumns = 1
        Else
            brRows = UBound(ref2B, 1) - LBound(ref2B, 1) + 1
            brColumns = UBound(ref2B, 2) - LBound(ref2B, 2) + 1
        End If
    End If
    
    If arColumns <> brColumns Then
        Debug.Print "Error: the swap sizes are different."
        Exit Function
    Else
        For i = 1 To WorksheetFunction.Min(aRows, bRows)
            For j = 1 To arColumns
                a(i, ref2A(j)) = b(i, ref2B(j))
            Next j
        Next i
    End If
    
    sReplaceColumns = a
            

End Function

Function IsColumn(ByVal testArray As Variant) As Boolean
'' Check if input array/range is a column vector
'
' Author: Yeol C. Seong
' Date: 2000/08/01
' Contact: monkeyquant@gmail.com
' REVISED:
'
    Dim nRow As Long, nCol As Long
    Dim tA As Variant
    IsColumn = False

    If TypeName(testArray) = "Range" Then
        tA = testArray.Value
        Set testArray = Nothing
        If Not IsArray(tA) Then MsgBox ("This is not array"): Exit Function
        
        nRow = UBound(tA, 2)
        nCol = UBound(tA, 1)
    Else
        tA = testArray
        testArray = Empty
        If Not IsArray(tA) Then MsgBox ("This is not array"): Exit Function
        
        nRow = UBound(tA, 2)
        nCol = UBound(tA, 1)

    End If
    
    If nCol = 1 And nRow > 1 Then IsColumn = True

End Function




Function IsRow(ByVal testArray As Variant) As Boolean
'' Check if input array/range is a row vector
'
' Author: Yeol C. Seong
' Date: 2000/08/01
' Contact: monkeyquant@gmail.com
' REVISED:
'
    Dim nRow As Long, nCol As Long
    Dim tA As Variant
    IsRow = False

    If TypeName(testArray) = "Range" Then
        tA = testArray.Value
        Set cA = Nothing
        If Not IsArray(tA) Then MsgBox ("This is not array"): Exit Function
        
        nRow = UBound(tA, 2)
        nCol = UBound(tA, 1)

    Else
        tA = testArray
        testArray = Empty
        If Not IsArray(tA) Then MsgBox ("This is not array"): Exit Function
        
        nRow = UBound(tA, 2)
        nCol = UBound(tA, 1)

    End If
    
    If nCol > 1 And nRow = 1 Then IsRow = True

End Function




Function IsMatrix(ByVal testArray As Variant) As Boolean
'' Check if input array/range is a matrix vector
'
' Author: Yeol C. Seong
' Date: 2000/08/01
' Contact: monkeyquant@gmail.com
' REVISED:
'
    Dim nRow As Long, nCol As Long
    Dim tA As Variant
    IsMatrix = False

    If TypeName(testArray) = "Range" Then
        tA = testArray.Value
        Set cA = Nothing
        If Not IsArray(tA) Then MsgBox ("This is not array"): Exit Function
        
        nRow = UBound(tA, 1)
        nCol = UBound(tA, 2)
    Else
        tA = cA
        testArray = Empty
        If Not IsArray(tA) Then MsgBox ("This is not array"): Exit Function
       
        nRow = UBound(tA, 1) - LBound(tA, 1) + 1
        nCol = UBound(tA, 2) - LBound(tA, 1) + 1

    End If
    
    If nCol > 1 And nRow > 1 Then IsMatrix = True

End Function




Function IsScalar(ByVal cA As Variant) As Boolean
'' Check if input array/range is a scalar vector
'
' Author: Yeol C. Seong
' Date: 2000/08/01
' Contact: monkeyquant@gmail.com
' REVISED:
'
    Dim nRow As Long, nCol As Long
    Dim tA As Variant

    If TypeName(cA) = "Range" Then
        tA = cA.Value
        Set cA = Nothing
        nRow = UBound(tA, 1)
        nCol = UBound(tA, 2)
        If Not IsArray(tA) Or (nCol = 1 And nRow = 1) Then
            IsScalar = True
        Else
            IsScalar = False
        End If
    Else
        tA = cA
        cA = Empty
        
        nRow = UBound(tA, 1) - LBound(tA, 1) + 1
        nCol = UBound(tA, 2) - LBound(tA, 2) + 1
        
        If Not IsArray(tA) Or (nCol = 1 And nRow = 1) Then
            IsScalar = True
        Else
            
            IsScalar = False
        End If
    End If

End Function




Public Function IsArrayEmpty(ByVal cArr As Variant) As Boolean
    On Error GoTo IS_EMPTY
    If (UBound(cArr) >= 0) Then Exit Function
IS_EMPTY:
        IsArrayEmpty = True
End Function




Public Function matMAX(ByVal cA As Variant, Optional ByVal cKey As Variant = "All") As Variant
' matMAX: Calculate Matrix Sum of columns, row, or all.
'
' For Example, A={1,2; 3,4}, Then matMAX(A, "row") = {3, 4} or matMAX(A, "column") ={2; 4}.
' Author: Yeol C. Seong
' Date: 2015/02/23
' Contact: yseong@uchicago.edu
' REVISED:
'           - Add row and column index of maximum value, not implemneted yet
'
    Dim nRowA As Variant, nColA As Variant
    Dim MAT_A As Variant
    Dim i As Long, j As Long, k As Long
    Dim cAns() As Variant
    Dim tDummy As Double
    Dim tKey As Variant

    If TypeName(cA) = "Range" Then
        MAT_A = cA.Value
        'Set cA = Nothing
        nRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
        nColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(cA) = 0 Then
            MAT_A = cA
            'Erase cA
            
            nRowA = UBound(MAT_A, 1) - LBound(MAT_A, 1) + 1       ' Measuring The length of the rows of a Matrix
            nColA = UBound(MAT_A, 2) - LBound(MAT_A, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_A(1 To nRowA, 1 To nColA)
        Else
            MAT_A = cA
            'Erase cA
            
            nRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
            nColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If
    
    If TypeName(cKey) = "Range" Then
        tKey = cKey.Value
        Set cKey = Nothing
    Else
        tKey = cKey
    End If
    If Not IsNumeric(tKey) Then
        If LCase(Left(tKey, 1)) = "r" Then
            tKey = 2
        ElseIf LCase(Left(tKey, 1)) = "c" Then
            tKey = 1
        Else
            tKey = 0
        End If
    End If
    
    Select Case tKey
        Case 1                                          ' Column Max
            ReDim cAns(1 To 1, 1 To nColA) As Variant
    
            For j = 1 To nColA
                '' Initiate the first value
                If nRowA Mod 2 = 1 Then
                    tDummy = MAT_A(1, j)
                    i = 2
                Else
                    If MAT_A(1, j) < MAT_A(2, j) Then
                        tDummy = MAT_A(2, j)
                    Else
                        tDummy = MAT_A(1, j)
                    End If
                    i = 3
                End If
                
                Do While (i < nRowA)
                ''For i = 3 To nRowA
                    '' For k = i + 1 To nRowA
                        If (MAT_A(i, j) < MAT_A(i + 1, j)) Then
                            If (MAT_A(i + 1, j) > tDummy) Then
                                tDummy = MAT_A(i + 1, j)
                            End If
                        Else
                            If (MAT_A(i, j) > tDummy) Then
                                tDummy = MAT_A(i, j)
                            End If
                        End If
                        
                        i = i + 2
                    '' Next k
                ''Next i
                Loop
                
                cAns(1, j) = tDummy
            
            Next j
            
            matMAX = cAns
        
        Case 2                                          ' Row Max
            ReDim cAns(1 To nRowA, 1 To 1) As Variant

            For i = 1 To nRowA
                '' Initiate the first value
                If nColA Mod 2 = 1 Then
                    tDummy = MAT_A(i, 1)
                    j = 2
                Else
                    If MAT_A(i, 1) < MAT_A(i, 2) Then
                        tDummy = MAT_A(i, 2)
                    Else
                        tDummy = MAT_A(i, 1)
                    End If
                    j = 3
                End If
                
                Do While (j < nColA)
                '' For j = 1 To nColA - 1
                    '' For k = j + 1 To nColA
                        If MAT_A(i, j) < MAT_A(i, j + 1) Then
                            If MAT_A(i, j + 1) > tDummy Then
                                tDummy = MAT_A(i, j + 1)
                            End If
                        Else
                            If MAT_A(i, j) > tDummy Then
                                tDummy = MAT_A(i, j)
                            End If
                        End If
                        
                        j = j + 2
                    '' Next k
                '' Next j
                Loop
                
                cAns(i, 1) = tDummy
            
            Next i
            
            matMAX = cAns
       
        Case Else
            
            matMAX = Application.WorksheetFunction.Max(MAT_A)
    
    End Select
    

End Function


Public Function matMIN(ByVal cA As Variant, Optional ByVal cKey As Variant = "All") As Variant
'
' sRowColumnSum               : Calculate Matrix Sum of columns or rows.
'
' For Example, A={1,2; 3,4}, Then Ans = {3, 7} or {4; 6}.

' %%Author: Yeol C. Seong
' %%Date: 2015/02/23
' %%Contact: yseong@uchicago.edu
' %%REVISED:
'
    Dim nRowA As Variant, nColA As Variant
    Dim MAT_A As Variant
    Dim i As Long, j As Long, k As Long
    Dim cAns() As Variant
    Dim tDummy As Double, tDum1 As Double
    Dim tKey As Variant

    If TypeName(cA) = "Range" Then
        MAT_A = cA.Value
        'Set cA = Nothing
        nRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
        nColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(cA) = 0 Then
            MAT_A = cA
            'Erase cA
            
            nRowA = UBound(MAT_A, 1) - LBound(MAT_A, 1) + 1       ' Measuring The length of the rows of a Matrix
            nColA = UBound(MAT_A, 2) - LBound(MAT_A, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_A(1 To nRowA, 1 To nColA)
        Else
            MAT_A = cA
            'Erase cA
            
            nRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
            nColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If
    
    If TypeName(cKey) = "Range" Then
        tKey = cKey.Value
        Set cKey = Nothing
    Else
        tKey = cKey
    End If
    If Not IsNumeric(tKey) Then
        If LCase(Left(tKey, 1)) = "r" Then
            tKey = 2
        ElseIf LCase(Left(tKey, 1)) = "c" Then
            tKey = 1
        Else
            tKey = 0
        End If
    End If
    
    Select Case tKey
    
        Case 1                                          ' Column Min
            ReDim cAns(1 To 1, 1 To nColA) As Variant
    
            For j = 1 To nColA
                '' Initiate the first value
                If nRowA Mod 2 = 1 Then
                    tDummy = MAT_A(1, j)
                    i = 2
                Else
                    If MAT_A(1, j) > MAT_A(2, j) Then
                        tDummy = MAT_A(2, j)
                    Else
                        tDummy = MAT_A(1, j)
                    End If
                    i = 3
                End If
                
                Do While (i < nRowA)
                ''For i = 1 To NoRowA - 1
                    ''For k = i + 1 To NoRowA
                        If MAT_A(i, j) > MAT_A(i + 1, j) Then
                            If MAT_A(i + 1, j) < tDummy Then
                                tDummy = MAT_A(i + 1, j)
                            End If
                        Else
                            If MAT_A(i, j) < tDummy Then
                                tDummy = MAT_A(i, j)
                            End If
                        End If
                        
                        i = i + 2
                    ''Next k
                ''Next i
                Loop
                cAns(1, j) = tDummy
            Next j

            matMIN = cAns
        
        Case 2                                          ' Row Min
            ReDim cAns(1 To nRowA, 1 To 1) As Variant

            For i = 1 To nRowA
                '' Initiate the first value
                If nColA Mod 2 = 1 Then
                    tDummy = MAT_A(i, 1)
                    j = 2
                Else
                    If MAT_A(i, 1) > MAT_A(i, 2) Then
                        tDummy = MAT_A(i, 2)
                    Else
                        tDummy = MAT_A(i, 1)
                    End If
                    j = 3
                End If
                
                Do While (j < nColA)

                '' For j = 1 To NoColA - 1
                    '' For k = j + 1 To NoColA
                        If MAT_A(i, j) > MAT_A(i, j + 1) Then
                            If MAT_A(i, j + 1) < tDummy Then
                                tDummy = MAT_A(i, j + 1)
                            End If
                        Else
                            If MAT_A(i, j) < tDummy Then
                                tDummy = MAT_A(i, j)
                            End If
                        End If
                        
                        j = j + 2
                    ''Next k
                ''Next j
                Loop
                cAns(i, 1) = tDummy
            Next i
                 
            matMIN = cAns
      
        Case Else
            
            matMIN = Application.WorksheetFunction.Min(MAT_A)
        
    End Select
    

End Function


Public Function matSUM(ByVal inputArray As Variant, Optional ByVal cKey As Variant = "All") As Variant
'
' sRowColumnSum               : Calculate Matrix Sum of columns or rows.
'
' For Example, A={1,2; 3,4}, Then Ans = {3, 7} or {4; 6}.

' %%Author: Yeol C. Seong
' %%Date: 2015/02/23
' %%Contact: yseong@uchicago.edu
' %%REVISED:
'

    Dim NoRowA As Variant, NoColA As Variant
''    Dim NoRowB As Variant, NoColB As Variant
    Dim MAT_A As Variant
''    Dim MAT_B As Variant
    Dim i As Long, j As Long, k As Long
    Dim cAns() As Variant
    Dim tDummy As Double
    Dim tKey As Variant

    If TypeName(inputArray) = "Range" Then
        MAT_A = inputArray.Value
        Set inputArray = Nothing
        
        NoRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
        NoColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
    Else
        If LBound(inputArray) = 0 Then
            MAT_A = inputArray
            inputArray = Empty
            
            NoRowA = UBound(MAT_A, 1) - LBound(MAT_A, 1) + 1       ' Measuring The length of the rows of a Matrix
            NoColA = UBound(MAT_A, 2) - LBound(MAT_A, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_A(1 To NoRowA, 1 To NoColA)
        Else
            MAT_A = inputArray
            inputArray = Empty
            
            NoRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
            NoColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If

    If TypeName(cKey) = "Range" Then
        tKey = cKey.Value
        Set cKey = Nothing
    Else
        tKey = cKey
        cKey = Empty
    End If
    
    If Not IsNumeric(tKey) Then
        If LCase(Left(tKey, 1)) = "r" Then
            tKey = 2
        ElseIf LCase(Left(tKey, 1)) = "c" Then
            tKey = 1
        Else
            tKey = 0
        End If
    End If

    Select Case tKey
        Case 1
            ReDim cAns(1 To 1, 1 To NoColA) As Variant
    
            For j = 1 To NoColA
                tDummy = 0
                For i = 1 To NoRowA
                    tDummy = tDummy + MAT_A(i, j)
                Next i
                cAns(1, j) = tDummy
            Next j
            matSUM = cAns
          
        Case 2
            ReDim cAns(1 To NoRowA, 1 To 1) As Variant
    
            For i = 1 To NoRowA
                tDummy = 0
                For j = 1 To NoColA
                    tDummy = tDummy + MAT_A(i, j)
                Next j
                    cAns(i, 1) = tDummy
            Next i
            matSUM = cAns
            
        Case Else
            matSUM = Application.WorksheetFunction.Sum(MAT_A)
      End Select


End Function



'
''
Public Function matPLUS(ByVal cA As Variant, ByVal cB As Variant) As Variant
'
' sMATSUM               : Calculate Matrix Sum of element by element.
'
' For Example, A={1,2; 3,4}, and B = {6,7;8,9}, Then Ans = {7,9; 11,13}.

' %%Author: Yeol C. Seong
' %%Date: 2000/08/01
' %%Contact: yseong@uchicago.edu
' %%REVISED:
'

    Dim NoRowA As Variant
    Dim NoColA As Variant
    Dim NoRowB As Variant
    Dim NoColB As Variant
    ''Dim MAT_A() As Variant
    Dim MAT_A As Variant
    ''Dim MAT_B() As Variant
    Dim MAT_B As Variant
    Dim i As Integer
    Dim j As Integer
    Dim cAns() As Variant
    
    ''' Very Important Function Procedure When we use Range Array as Input.
    ''MAT_A = A.Value
    ''MAT_B = b.Value
    
    ''NoRowA = UBound(MAT_A, 1)                           ' Measuring The length of the rows of a Matrix
    ''NoColA = UBound(Application.Transpose(MAT_A), 1)    ' Measuring The length of the size of columns of a Matrix
    ''NoRowB = UBound(MAT_B, 1)                           ' Measuring The length of the rows of a Matrix
    ''NoColB = UBound(Application.Transpose(MAT_B), 1)    ' Measuring The length of the size of columns of a Matrix

    If TypeName(cA) = "Range" Then
        MAT_A = cA.Value
        'Set cA = Nothing
        NoRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
        NoColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(cA) = 0 Then
            MAT_A = cA
            'Erase cA
            
            NoRowA = UBound(MAT_A, 1) - LBound(MAT_A, 1) + 1       ' Measuring The length of the rows of a Matrix
            NoColA = UBound(MAT_A, 2) - LBound(MAT_A, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_A(1 To NoRowA, 1 To NoColA)
        Else
            MAT_A = cA
            'Erase cA
            
            NoRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
            NoColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If
    
    If TypeName(cB) = "Range" Then
        MAT_B = cB.Value
        'Set cB = Nothing
        NoRowB = UBound(MAT_B, 1)       ' Measuring The length of the rows of a Matrix
        NoColB = UBound(MAT_B, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(cB) = 0 Then
            MAT_B = cB
            'Erase cB
            
            NoRowB = UBound(MAT_B, 1) - LBound(MAT_B, 1) + 1       ' Measuring The length of the rows of a Matrix
            NoColB = UBound(MAT_B, 2) - LBound(MAT_B, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_A(1 To NoRowB, 1 To NoColB)
        Else
            MAT_B = cB
            'Erase cB
            
            NoRowB = UBound(MAT_B, 1)       ' Measuring The length of the rows of a Matrix
            NoColB = UBound(MAT_B, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If
'

    If NoRowA = NoRowB Then
        If NoColA = NoColB Then
            ReDim Preserve MAT_A(1 To NoRowA, 1 To NoColA)
            ReDim Preserve MAT_B(1 To NoRowB, 1 To NoColB)
            ReDim cAns(1 To NoRowA, 1 To NoColA) As Variant
    
            For i = 1 To NoRowA
                For j = 1 To NoColA
                    cAns(i, j) = MAT_A(i, j) + MAT_B(i, j)
                Next j
            Next i
        matPLUS = cAns
        Else
            MsgBox "The size of matrix is not matched"
        End If
    Else
        MsgBox "The size of matrix is not matched"
    End If
End Function

 
 
Public Function matMINUS(ByVal cA As Variant, ByVal cB As Variant) As Variant
'
' sMATSUM               : Calculate Matrix Sum of element by element.
'
' For Example, A={1,2; 3,4}, and B = {6,7;8,9}, Then Ans = {7,9; 11,13}.

' %%Author: Yeol C. Seong
' %%Date: 2000/08/01
' %%Contact: yseong@uchicago.edu
' %%REVISED:
'

    Dim NoRowA As Variant
    Dim NoColA As Variant
    Dim NoRowB As Variant
    Dim NoColB As Variant
    ''Dim MAT_A() As Variant
    Dim MAT_A As Variant
    ''Dim MAT_B() As Variant
    Dim MAT_B As Variant
    Dim i As Integer
    Dim j As Integer
    Dim cAns() As Variant
    
    ''' Very Important Function Procedure When we use Range Array as Input.
    ''MAT_A = A.Value
    ''MAT_B = b.Value
    
    ''NoRowA = UBound(MAT_A, 1)                           ' Measuring The length of the rows of a Matrix
    ''NoColA = UBound(Application.Transpose(MAT_A), 1)    ' Measuring The length of the size of columns of a Matrix
    ''NoRowB = UBound(MAT_B, 1)                           ' Measuring The length of the rows of a Matrix
    ''NoColB = UBound(Application.Transpose(MAT_B), 1)    ' Measuring The length of the size of columns of a Matrix

    If TypeName(cA) = "Range" Then
        MAT_A = cA.Value
        'Set cA = Nothing
        NoRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
        NoColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(cA) = 0 Then
            MAT_A = cA
            'Erase cA
            
            NoRowA = UBound(MAT_A, 1) - LBound(MAT_A, 1) + 1       ' Measuring The length of the rows of a Matrix
            NoColA = UBound(MAT_A, 2) - LBound(MAT_A, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_A(1 To NoRowA, 1 To NoColA)
        Else
            MAT_A = cA
            'Erase cA
            
            NoRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
            NoColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If
    
    If TypeName(cB) = "Range" Then
        MAT_B = cB.Value
        'Set cB = Nothing
        NoRowB = UBound(MAT_B, 1)       ' Measuring The length of the rows of a Matrix
        NoColB = UBound(MAT_B, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(cB) = 0 Then
            MAT_B = cB
            'Erase cB
            
            NoRowB = UBound(MAT_B, 1) - LBound(MAT_B, 1) + 1       ' Measuring The length of the rows of a Matrix
            NoColB = UBound(MAT_B, 2) - LBound(MAT_B, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_A(1 To NoRowB, 1 To NoColB)
        Else
            MAT_B = cB
            'Erase cB
            
            NoRowB = UBound(MAT_B, 1)       ' Measuring The length of the rows of a Matrix
            NoColB = UBound(MAT_B, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If
'

    If NoRowA = NoRowB Then
        If NoColA = NoColB Then
            ReDim Preserve MAT_A(1 To NoRowA, 1 To NoColA)
            ReDim Preserve MAT_B(1 To NoRowB, 1 To NoColB)
            ReDim cAns(1 To NoRowA, 1 To NoColA) As Variant
    
            For i = 1 To NoRowA
                For j = 1 To NoColA
                    cAns(i, j) = MAT_A(i, j) - MAT_B(i, j)
                Next j
            Next i
        matMINUS = cAns
        Else
            MsgBox "The size of matrix is not matched"
        End If
    Else
        MsgBox "The size of matrix is not matched"
    End If
End Function
 
 
 
 
Public Function matZEROS(ByVal cA As Variant, Optional ByVal cB As Variant) As Variant
'
' sMATZEROS             : Make ZERO Matrix like zeros in MATLAB.
'
' For Example, A=2, and B = 3, Then Ans = {0, 0, 0; 0, 0, 0}.
'
' %%Author: Yeol C. Seong
' %%Date: 2000/08/01
' %%Contact: yseong@uchicago.edu
' %%REVISED:
'
    Dim SCAL_A As Long, SCAL_B As Long
    Dim i As Integer, j As Integer
    Dim cAns() As Variant
    

    If TypeName(cA) = "Range" Then
        SCAL_A = cA.Value
        Set cA = Nothing
    Else
        SCAL_A = cA
        cA = Empty
    End If

    
    If IsMissing(cB) Then
        SCAL_B = SCAL_A
    Else
            
        If TypeName(cB) = "Range" Then
            SCAL_B = cB.Value
            Set cB = Nothing
        Else
            SCAL_B = cB
            cB = Empty
        End If
    End If

    
    ' Very Important Function Procedure When we use Array as Input.
    ReDim cAns(1 To SCAL_A, 1 To SCAL_B)
    For i = 1 To SCAL_A
        For j = 1 To SCAL_B
            cAns(i, j) = 0
        Next j
    Next i
        matZEROS = cAns
End Function

 

Public Function matONES(ByVal rowNum As Variant, Optional ByVal colNum As Variant) As Variant

'Attribute sMATONES.VB_Description = "Create M by N Matrix with element 1."
'
' sMATZEROS             : Make ZERO Matrix like zeros in MATLAB.
'
' For Example, A=2, and B = 3, Then Ans = {0, 0, 0; 0, 0, 0}.
'
' %%Author: Yeol C. Seong
' %%Date: 2000/08/01
' %%Contact: yseong@uchicago.edu
' %%REVISED:
'
    
    Dim SCAL_A As Long, SCAL_B As Long
    Dim i As Integer, j As Integer
    Dim cAns() As Variant
    
    If TypeName(rowNum) = "Range" Then
        If rowNum.Value = 0 Then
            MsgBox "Please enter greater than zero value"
            Exit Function
        End If

        SCAL_A = rowNum.Value
        Set rowNum = Nothing
    Else
        If rowNum = 0 Then
            MsgBox "Please enter greater than zero value"
            Exit Function
        End If

        SCAL_A = rowNum
        rowNum = Empty
    End If
    
        
    If IsMissing(colNum) Then
        SCAL_B = SCAL_A
    Else
        If TypeName(colNum) = "Range" Then
            SCAL_B = colNum.Value
            Set colNum = Nothing
        Else
            SCAL_B = colNum
            colNum = Empty
        End If
    End If
    
    ' Very Important Function Procedure When we use Array as Input.
    ReDim Preserve cAns(1 To SCAL_A, 1 To SCAL_B)
    
    For i = 1 To SCAL_A
        For j = 1 To SCAL_B
            cAns(i, j) = 1
        Next j
    Next i
    
        matONES = cAns

End Function

 
Public Function matIDENT(ByVal rowNum As Variant, Optional ByVal colNum As Variant, Optional cKey As Integer = 0) As Variant
' sMATZEROS: Make identical Matrix like eyes in MATLAB.
'
' For Example, A=2, and B = 3, Then Ans = {0, 0, 0; 0, 0, 0}.
'
' %%Author: Yeol C. Seong
' %%Date: 2000/08/01
' %%Contact: yseong@uchicago.edu
' %%REVISED:
'
    Dim i As Integer, j As Integer, SCAL_A As Integer, SCAL_B As Integer
    Dim cAns() As Variant
    Dim tAns As VbMsgBoxResult

    If TypeName(rowNum) = "Range" Then
        SCAL_A = rowNum.Value
        Set rowNum = Nothing
    Else
        SCAL_A = rowNum
        rowNum = Empty
    End If
    
    If IsMissing(colNum) Then
        SCAL_B = SCAL_A
    Else
        If TypeName(colNum) = "Range" Then
            SCAL_B = colNum.Value
            Set colNum = Nothing
        Else
            SCAL_B = colNum
            colNum = Empty
        End If
    End If
    
    If SCAL_A <> SCAL_B Then
        
        Select Case cKey
            Case -1
                SCAL_A = WorksheetFunction.Min(SCAL_A, SCAL_B)
                SCAL_B = SCAL_A
                GoTo 100
        
            Case 1
                SCAL_A = WorksheetFunction.Max(SCAL_A, SCAL_B)
                SCAL_B = SCAL_A
                GoTo 100
                
            Case 0
                SCAL_B = SCAL_A
                GoTo 100
            Case Else
                GoTo 100
        End Select
    Else
        GoTo 100
    End If
    
    
    ' Very Important Function Procedure When we use Array as Input.
100    ReDim cAns(1 To SCAL_A, 1 To SCAL_B)
    
    For i = 1 To SCAL_A
        For j = 1 To SCAL_B
            If i = j Then
                cAns(i, j) = 1
            Else
                cAns(i, j) = 0
            End If
        Next j
    Next i
    matIDENT = cAns
End Function




Public Function matDiag(ByVal rowNum As Variant, Optional ByVal colNum As Variant, Optional ByVal diagNum As Variant = 1, Optional ByVal cKey As Integer = 0) As Variant
' matDiag: Make doagonal Matrix like diag in MATLAB.
'
' For Example, A=2, and B = 3, Then Ans = {d, 0, 0; 0, d, 0}.
'
' %%Author: Yeol C. Seong
' %%Date: 2000/08/01
' %%Contact: yseong@uchicago.edu
' %%REVISED:
'
    Dim i As Integer, j As Integer, SCAL_A As Integer, SCAL_B As Integer
    Dim cAns() As Variant
    Dim tAns As VbMsgBoxResult

    If TypeName(rowNum) = "Range" Then
        SCAL_A = rowNum.Value
        Set rowNum = Nothing
    Else
        SCAL_A = rowNum
        rowNum = Empty
    End If
    
    If IsMissing(colNum) Then
        SCAL_B = SCAL_A
    Else
        If TypeName(colNum) = "Range" Then
            SCAL_B = colNum.Value
            Set colNum = Nothing
        Else
            SCAL_B = colNum
            colNum = Empty
        End If
    End If
    
    If SCAL_A <> SCAL_B Then
        Select Case cKey
            Case -1
                SCAL_A = WorksheetFunction.Min(SCAL_A, SCAL_B)
                SCAL_B = SCAL_A
                GoTo 100
        
            Case 1
                SCAL_A = WorksheetFunction.Max(SCAL_A, SCAL_B)
                SCAL_B = SCAL_A
                GoTo 100
                
            Case 0
                SCAL_B = SCAL_A
                GoTo 100
            Case Else
                GoTo 100
        End Select
    Else
        GoTo 100
    End If
    
    
    ' Very Important Function Procedure When we use Array as Input.
100    ReDim cAns(1 To SCAL_A, 1 To SCAL_B)
    
    For i = 1 To SCAL_A
        For j = 1 To SCAL_B
            If i = j Then
                cAns(i, j) = diagNum
            Else
                cAns(i, j) = 0
            End If
        Next j
    Next i
    matDiag = cAns
End Function




Public Function matPROD(ByVal leftArr As Variant, ByVal rightArr As Variant) As Variant
'
' sMATPROD              : Calculate Matrix Multiplication of element by element.
'
'For Example, A={1,2; 3,4}, and B = {6,7;8,9}, Then Ans = {6,14; 24,36}.
'
'
' %%Author: Yeol C. Seong
' %%Date: 2001/08/01
' %%Contact: yseong@uchicago.edu or yeol.seong@bmo.com
' %%REVISED:
'
    
    Dim NoRowA As Variant, NoColA As Variant    ' size of left matrix
    Dim NoRowB As Variant, NoColB As Variant    ' size of right matrix
    Dim MAT_A As Variant, MAT_B As Variant      ' all arrrays including range into vba array
    Dim i As Integer, j As Integer, k As Integer    ' dummy variables
    Dim cAns() As Variant
    Dim tAns As Double                          ' column and row calculation
    
    ' to be used as a Worksheet Function and VBA Array together
    ' all arrays from Rangestarting with base 1, not 0
    If TypeName(leftArr) = "Range" Then
        MAT_A = leftArr.Value                ' When we use Array as Input.
        Set leftArr = Nothing                ' Free up memory
        NoRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
        NoColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(leftArr) = 0 Then
            MAT_A = leftArr
            leftArr = Empty
            NoRowA = UBound(MAT_A, 1) - LBound(MAT_A, 1) + 1       ' Measuring The length of the rows of a Matrix
            NoColA = UBound(MAT_A, 2) - LBound(MAT_A, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_A(1 To NoRowA, 1 To NoColA)
        Else
            MAT_A = leftArr
            leftArr = Empty
            NoRowA = UBound(MAT_A, 1)   ' Measuring The length of the rows of a Matrix
            NoColA = UBound(MAT_A, 2)   ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If
    
    ' to be used as a Worksheet Function and VBA Array together
    ' all arrays from Rangestarting with base 1, not 0
    If TypeName(rightArr) = "Range" Then
        MAT_B = rightArr.Value
        Set rightArr = Nothing
        NoRowB = UBound(MAT_B, 1)       ' Measuring The length of the rows of a Matrix
        NoColB = UBound(MAT_B, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(rightArr) = 0 Then
            MAT_B = rightArr
            rightArr = Empty
            NoRowB = UBound(MAT_B, 1) - LBound(MAT_B, 1) + 1       ' Measuring The length of the rows of a Matrix
            NoColB = UBound(MAT_B, 2) - LBound(MAT_B, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_B(1 To NoRowB, 1 To NoColB)
        Else
            MAT_B = rightArr
            rightArr = Empty
            NoRowB = UBound(MAT_B, 1)   ' Measuring The length of the rows of a Matrix
            NoColB = UBound(MAT_B, 2)   ' Measuring The length of the size of columns of a Matrix
        End If
    End If

' apply multiplication by checking if left Array's Column size the same as the row size of right Array
'    If NoRowA = NoRowB Then
'        If NoColA = NoColB Then
    If NoColA = NoRowB Then
'            ReDim Preserve MAT_A(1 To NoRowA, 1 To NoColA)
'            ReDim Preserve MAT_B(1 To NoRowB, 1 To NoColB)
        ReDim cAns(1 To NoRowA, 1 To NoColB) As Variant
        
        tAns = 0#
        
        For i = 1 To NoRowA
            For j = 1 To NoColB
                For k = 1 To NoColA
                    tAns = tAns + MAT_A(i, k) * MAT_B(k, j)
                Next k
                cAns(i, j) = tAns
                tAns = 0#
            Next j
        Next i
        
        matPROD = cAns
'        Else
'            MsgBox "The size of matrix is not matched"
'        End If
    Else
        MsgBox "Error 13: The size of matrix is not matched"
            
        matPROD = "Error 13"
    End If
    
End Function

 

Public Function matSCAL(ByVal inputArr As Variant, ByVal inputScalar As Variant) As Variant
'
' sMATSCALPROD              : Calculate Matrix Multiplication with a Scalar.
'
' For Example, A={1,2; 3,4}, and B= 2, Then Ans = {2,4; 6,8}.
'
'
' %%Author: Yeol C. Seong
' %%Date: 2001/08/01
' %%Contact: yseong@uchicago.edu or yeol.seong@bmo.com
' %%REVISED:
    '
     
    Dim NoRowA As Variant, NoColA As Variant
    Dim MAT_A As Variant, SCAL_B As Variant
    Dim i As Integer, j As Integer
    Dim cAns() As Variant
    
    ' Very Important Function Procedure When we use Array as Input.
''    MAT_A = A.Value
''    SCAL_B = b
    If TypeName(inputArr) = "Range" Then
        MAT_A = inputArr.Value
        Set inputArr = Nothing
        NoRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
        NoColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(inputArr) = 0 Then
            MAT_A = inputArr
            cA = Empty
            
            NoRowA = UBound(MAT_A, 1) - LBound(MAT_A, 1) + 1       ' Measuring The length of the rows of a Matrix
            NoColA = UBound(MAT_A, 2) - LBound(MAT_A, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve MAT_A(1 To NoRowA, 1 To NoColA)
        Else
            MAT_A = inputArr
            cA = Empty
            
            NoRowA = UBound(MAT_A, 1)       ' Measuring The length of the rows of a Matrix
            NoColA = UBound(MAT_A, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If
    
    If TypeName(inputScalar) = "Range" Then
        SCAL_B = inputScalar.Value
        Set inputScalar = Nothing
    Else
        SCAL_B = inputScalar
        inputScalar = Empty
    End If
    
    
    ReDim cAns(1 To NoRowA, 1 To NoColA) As Variant
    
    For i = 1 To NoRowA
        For j = 1 To NoColA
            cAns(i, j) = MAT_A(i, j) * SCAL_B
        Next j
    Next i
    
    matSCAL = cAns

End Function

'
'''
''
Public Function LookupByVal(ByVal LookupValue As Variant, ByVal rngSearch As Variant, ByVal ref_index_num As Variant, _
    Optional ByVal search_index_num As Variant = 1, Optional ByVal HorizontalMatch As Boolean = False) As Variant
' %% sLOOKUP replaces Vlookup and HLookup and over its limit
' %% Input:
'           LookupValue: a single entry
'           ref_index_num: target array to refer
'           rngSearch: a matrix
'           search_index_num: a column or row to look up the lookupvalue
'
' %%Author: Yeol C. Seong
' %%Date: 2000/08/01
' %%Contact: yseong@uchicago.edu
' %%REVISED:
'

    'Application.Volatile

    If HorizontalMatch = True Then
    ' %% Hlookup
        LooupByVal = Application.Index(rngSearch, ref_index_num, Application.Match(LookupValue, _
            Application.Index(Application.Transpose(rngSearch), 0, search_index_num), 0))
    Else
    ' %% Vlookup
        LookupByVal = Application.Index(rngSearch, Application.Match(LookupValue, _
            Application.Index(rngSearch, 0, search_index_num), 0), ref_index_num)
    End If

End Function

'
''
Public Function LookupByRef(ByRef LookupValue As Variant, ByRef rngSearch As Range, ByRef ref_index_num As Variant, _
    Optional ByRef search_index_num As Variant = 1, Optional ByRef HorizontalMatch As Boolean = False) As Variant
' %% sLOOKUP replaces Vlookup and HLookup and over its limit
' %% Input:
'           LookupValue: a single entry
'           ref_index_num: target array to refer
'           rngSearch: a matrix
'           search_index_num: a column or row to look up the lookupvalue
'
' %%Author: Yeol C. Seong
' %%Date: 2000/08/01
' %%Contact: yseong@uchicago.edu
' %%REVISED:
'

    Application.Volatile

    If HorizontalMatch = True Then
    ' %% Hlookup
        LookupByRef = Application.Index(rngSearch, ref_index_num, Application.Match(LookupValue, _
            Application.Index(Application.Transpose(rngSearch), 0, search_index_num), 0))
    Else
    ' %% Vlookup
        LookupByRef = Application.Index(rngSearch, Application.Match(LookupValue, _
            Application.Index(rngSearch, 0, search_index_num), 0), ref_index_num)
    End If

End Function



Function sPolyfit(ByVal cX As Variant, ByVal cY As Variant, Optional ByVal cOrder As Variant = 1, Optional ByVal cOpt As Variant = "Coefficient") As Variant
' Polyfit each column of a matrix Y as a function of an input vector Xand return 1 Column vector of coeffients
' to a matrix P. The lower index corresponds to higher order of the the polynomial.
'
'   Err:            Frobenius norm of the estimation error for each column.
'   Adapted from the polyfit function of Mathworks
'
'       Example 1:
'       Fit a polynomial p of degree 1 to the (x,y) data:
'    %       x = 1:50;
'    %       y = -0.3*x + 2*randn(1,50);
'    %       p = colPolyFit(x, y, 1);
'    %
'    % Example 2:
'    %       % Fit multiple polynomial where different columns of Y will
'    %       correspond to different polynomial
'    %       x = 1:50;
'    %       y1 = -0.3*x + 2*randn(1,50);
'    %       y2 = 2*x - 5*randn(1,50);
'    %       y = [y1(:) y2(:)];
'    %       p = colPolyFit(x, y, 1);
'    %
'    %       Evaluation of the fitted result
'    %       y1 = colPolyVal(p, x);
'    % Tan H. Nguyen - MIT, EECS 2018
'    % thnguyn@mit.edu
'
'   Author: Yeol Seong,  Beyond Financial Group
'   Date: 4/30/2022
'

    Dim nRowX As Long, nRowY As Long, nColX As Long, nColY As Long, nRowZ As Long, nColZ As Long
    Dim nLength As Long
    Dim p As Variant, Err As Variant, tOpt As Variant
    
    Dim tX As Variant, tY As Variant, tOrder As Variant
    Dim i As Long, j As Long, tSize As Long

    Dim cAns() As Variant
    
    If TypeName(cX) = "Range" Then
        tX = cX.Value
        'Set cX = Nothing
        nRowX = UBound(tX, 1)       ' Measuring The length of the rows of a Matrix
        nColX = UBound(tX, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(cX) = 0 Then
            tX = cX
            'Erase cX
            
            nRowX = UBound(tX, 1) - LBound(tX, 1) + 1       ' Measuring The length of the rows of a Matrix
            nColX = UBound(tX, 2) - LBound(tX, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve tX(1 To nRowX, 1 To nColX)
        Else
            tX = cX
            'Erase cX
            
            nRowX = UBound(tX, 1)       ' Measuring The length of the rows of a Matrix
            nColX = UBound(tX, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
    End If
    
    If TypeName(cY) = "Range" Then
        tY = cY.Value
        'Set cY = Nothing
        nRowY = UBound(tY, 1)       ' Measuring The length of the rows of a Matrix
        nColY = UBound(tY, 2)       ' Measuring The length of the size of columns of a Matrix
        
    Else
        If LBound(cY) = 0 Then
            tY = cY
            'Erase cY
            
            nRowY = UBound(tY, 1) - LBound(tY, 1) + 1       ' Measuring The length of the rows of a Matrix
            nColY = UBound(tY, 2) - LBound(tY, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve tY(1 To nRowY, 1 To nColY)
        Else
            tY = cY
            'Erase cY
            
            nRowY = UBound(tY, 1)       ' Measuring The length of the rows of a Matrix
            nColY = UBound(tY, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
    End If

    
    If TypeName(cOrder) = "Range" Then
        tOrder = cOrder.Value
        'Set cOrder = Nothing
        
        If IsArray(tOrder) Then
            If IsMatrix(tOrder) Then
                tOrder = sArray2Vector(tOrder, False)
            End If
''            nRowZ = UBound(tOrder, 1)
            nColZ = UBound(tOrder, 2)
        Else
''            nRowZ = 1
            nColZ = 1
        End If
        
    Else
''        tOrder = cOrder
        If IsArray(cOrder) Then
            If IsMatrix(cOrder) Then
                tOrder = sArray2Vector(cOrder, False)
            End If
            If LBound(tOrder) = 0 Then
''                nRowZ = UBound(tOrder) - LBound(tOrder) + 1
                nColZ = UBound(tOrder, 2) - LBound(tOrder, 1) + 1
''
''                ReDim Preserve tOrder(1 To nRowZ, 1 To nColZ)
                
            Else
''                nRowZ = UBound(tOrder, 1)
                nColZ = UBound(tOrder, 2)
            End If
        Else
            tOrder = cOrder
''            nRowZ = 1
            nColZ = 1
        End If
    End If
    
''    nLength = Application.WorksheetFunction.Max(nRowZ, nColZ)
    
    If TypeName(cOpt) = "Range" Then
        tOpt = cOpt.Value
        Set cOpt = Nothing
    Else
        tOpt = cOpt
    End If
    
    If Not IsNumeric(tOpt) Then
        Select Case LCase(Left(tOpt, 1))
            Case "p", "c"
                tOpt = 1
            Case "e"
                tOpt = 2
            Case Else
                tOpt = 3
        End Select
    End If
    
    
    
    ' Check the input type and size
    
    If IsScalar(tX) Or IsScalar(tY) Then
        MsgBox ("Input should be either matrix or vector.")
        Exit Function
    ElseIf IsMatrix(tX) Or IsMatrix(tY) Then
        If (IsMatrix(tX) And Not IsMatrix(tY)) Or (Not IsMatrix(tX) And IsMatrix(tY)) Then
            MsgBox ("Out of Scope: the X and Y are different size")
            Exit Function
        Else
            If (sColumnSize(tX) <> sColumnSize(tY)) Or (sRowSize(tX) <> sRowSize(tY)) Then
                MsgBox ("Out of Scope: the X and Y are different size")
                Exit Function
''            Else
''                If IsMatrix(tOrder) Then
''                    tOrder = sArray2Vector(tOrder)
''                End If
            End If
        End If
    ElseIf IsColumn(tX) Then
        If IsColumn(tY) Then
            If UBound(tX, 2) <> UBound(tY, 2) Then
                MsgBox ("Out of Scope: the X and Y are different size")
                Exit Function
            End If
        ElseIf IsRow(tY) Then
            If UBound(tX, 2) <> UBound(tY, 1) Then
                MsgBox ("Out of Scope: the X and Y are different size")
                Exit Function
            Else
                tY = sResize(tY, nRowX, 1)
            End If
        End If
    ElseIf IsRow(tX) Then
        If IsRow(tY) Then
            If UBound(tX, 2) <> UBound(tY, 2) Then
                MsgBox ("Out of Scope: the X and Y are different size")
                Exit Function
            Else
                tSize = nColX
                nColX = nRowX
                nRowX = tSize
                
                tX = sResize(tX, nRowX, 1)
                tY = sResize(tY, nRowX, 1)
            End If
        ElseIf IsColumn(tY) Then
            If UBound(tX, 2) <> UBound(tY, 1) Then
                MsgBox ("Out of Scope: the X and Y are different size")
                Exit Function
            Else
                tSize = nColX
                nColX = nRowX
                nRowX = tSize
                tX = sResize(tX, nRowX, 1)
            End If
        End If
    Else
        MsgBox ("Out of Scope: The inputs for X and Y should have the same dimension.")
        Exit Function
    End If
        
    If nColZ > 1 Then
        If (nColX > nColZ) Then
            Dim tAppend As Variant
            tAppend = sMATSCAL(sMATONES(1, nColX - nColZ), tOrder(nColZ))
            tOrder = sAppendArray(tOrder, tAppend, "Row")
            tOrder = sResize(tOrder, 1, nColX)
        End If
    End If
                
    'ReDim cAns(1 To nRowX, 1 To nColX) As Variant
    
    Dim nCoefs As Integer
    nCoefs = tOrder + 1
    
    cAns = sMATONES(nRowX, nCoefs)
    
    For k = nCoefs To 1 Step -1
        For i = 1 To nRowX
            For j = 1 To nColX
                cAns(i, k) = cAns(i, k + 1) * tX(i, j)
            Next j
        Next i
    Next k
    
'    % Compute the Moore-Penrose inverse to solve for p in the coefficients VV * p = Y;
    Dim tAll As Variant
    Dim tQ As Variant, tQTY As Variant
    Dim tR As Variant, tRINV As Variant
'    Dim tP As Variant
    
'    tAll = sQRDEOMP(cAns, "a", "column")
    tQ = sQRDECOMP(cAns, "q")
    tQTY = Application.WorksheetFunction.MMult(Application.WorksheetFunction.Transpose(tQ), tY)
        
    tR = sQRDECOMP(cAns, "r")
    
    On Error GoTo ERRORHANDLE
    
    tRINV = Application.WorksheetFunction.MInverse(tR)
    p = Application.WorksheetFunction.MMult(tRINV, tQTY)       ' p = R\(Q'*Y);
        
    Dim tMu As Variant
    tMu = sMATMINUS(tY, Application.WorksheetFunction.MMult(cAns, p))
    tMu = sMATPROD(tMu, tMu)
    Err = Sqr(sMATSUM(tMu, "Column"))
    'err = sqrt(sum(mu .^2, 1))
    
    Select Case tOpt
        Case 1
            sPolyfit = p
        Case 2
            sPolyfit = Err
        Case Else
            sPolyfit = sAppendArray(p, Err)
    End Select
    
ERRORHANDLE:
    Exit Function

End Function

Function CHOL(matrix As Range)

Dim i As Integer, j As Integer, k As Integer, N As Integer
Dim a() As Double 'the original matrix
Dim Element As Double
Dim L_Lower() As Double

N = matrix.Columns.Count

ReDim a(1 To N, 1 To N)
ReDim L_Lower(1 To N, 1 To N)

For i = 1 To N
    For j = 1 To N
        a(i, j) = matrix(i, j).Value
        L_Lower(i, j) = 0
    Next j
Next i

For i = 1 To N
    For j = 1 To N
        Element = a(i, j)
        For k = 1 To i - 1
            Element = Element - L_Lower(i, k) * L_Lower(j, k)
        Next k
        If i = j Then
            L_Lower(i, i) = Sqr(Element)
        ElseIf i < j Then
            L_Lower(j, i) = Element / L_Lower(i, i)
        End If
    Next j
Next i

CHOL = Application.WorksheetFunction.Transpose(L_Lower)

End Function



Function sQRDECOMP(ByVal cA As Variant, Optional ByVal cKey As Variant = "All", Optional ByVal cDirect As Variant = "Column") As Variant
    Dim nRowA As Variant, nColA As Variant
    Dim Q As Variant, R As Variant, tKey As Variant
    Dim i As Long, j As Long, k As Long
    
    If TypeName(cA) = "Range" Then
        Q = cA.Value
        'Set cA = Nothing
        nRowA = UBound(Q, 1)       ' Measuring The length of the rows of a Matrix
        nColA = UBound(Q, 2)       ' Measuring The length of the size of columns of a Matrix
    Else
        If LBound(cA) = 0 Then
            Q = cA
            'Erase cA
            
            nRowA = UBound(Q, 1) - LBound(Q, 1) + 1       ' Measuring The length of the rows of a Matrix
            nColA = UBound(Q, 2) - LBound(Q, 2) + 1       ' Measuring The length of the size of columns of a Matrix
            
            ReDim Preserve Q(1 To nRowA, 1 To nColA)
        Else
            Q = cA
            'Erase cA
            
            nRowA = UBound(Q, 1)       ' Measuring The length of the rows of a Matrix
            nColA = UBound(Q, 2)       ' Measuring The length of the size of columns of a Matrix
        End If
        
    End If
    
    If TypeName(cKey) = "Range" Then
        tKey = cKey.Value
        'Set cKey = Nothing
    Else
        tKey = cKey
    End If
    
    If Not IsNumeric(tKey) Then
        Select Case LCase(Left(tKey, 1))
            Case "r", "u"
                tKey = 1
            Case "q", "l"
                tKey = 2
            Case Else
                tKey = 0
        End Select
    End If
 
        
'    Q = matZEROS(nRowA, nColA)                        'QR
'    ReDim Preserve Q(1 To nRowA, 1 To nColA)
    R = matZEROS(nRowA, nColA)                         'QR
'    ReDim Preserve R(1 To nRowA, 1 To nColA)
    
    
'' Gram Schmidt Process
    For j = 1 To nColA
        R(j, j) = 0
        For i = 1 To nRowA
            R(j, j) = R(j, j) + Q(i, j) ^ 2
        Next i
        
        If R(j, j) = 0 Then MsgBox ("A has linearly dependent columns"): Exit Function
        
        R(j, j) = Sqr(R(j, j))
        
        For i = 1 To nRowA
            Q(i, j) = Q(i, j) / R(j, j)
        Next i
        
        For k = (j + 1) To nColA
            
            R(j, k) = 0
            
            For i = 1 To nRowA
                R(j, k) = R(j, k) + Q(i, j) * Q(i, k)
            Next i
            
            For i = 1 To nRowA
                Q(i, k) = Q(i, k) - Q(i, j) * R(j, k)
            Next i
        Next k
    Next j
'
    Select Case tKey
        Case 0
            sQRDECOMP = prependArray(Q, R, cDirect)
        Case 1
            sQRDECOMP = R
        Case 2
            sQRDECOMP = Q
    End Select
            
End Function


Function PositionInArray(ByVal searchValue As Variant, ByVal targetRef As Variant, Optional ByVal idxRef As Long = 1, _
                         Optional ByVal matchType As Variant = 0, Optional ByVal searchType As Variant = "all", _
                         Optional ByVal resultKey As Integer) As Variant
' Returns the all, first, or last position of an element within any-dimension array. It returns 0 if the element is not in the array,
' and -1 if there is an error (Limited to 1D - 2D)
'
' Input
'   searchValue         Lookup Value(s) - either scalar or 1D array (not implemented yet)
'   targetRef           Reference Matrix(Array)
'   idxRef              column index to refer in targetRef. For the row, transpose targetRef
'   matchType           0 for exact match, -1 for ignore cases, and all other positive integers for number of characters from left
'                       numeric searchValue override for 0
'   searchType          "all" or 0 for all indices, "first" or 1 for the first matching index, and "last" for the last matching index
'                       any others for the nth matching idex
'   resultKey           1 for the row values, 2 for column values, and 0 for both



    Dim i, j, k, pos, rCount, cCount, searchKeyCount As Long
    Dim tItem As Variant
    Dim tAns, tArr As Variant                               ' result, reference array respectively
    Dim tSearch As Variant
    Dim searchKey As Integer                                ' numeric (0) or string (1), all other objects (3) which are error
    Dim searchBase, refBase As Integer                      ' Array Base
    Dim tPosition, nRow, nCol, nnRow, nnCol, nSearch As Long
    Dim tDirect As String                                   ' row or column
    
    On Error Resume Next
    
    '## searchValue
    '# Test Range Object for searchValue
    If TypeName(searchValue) = "Range" Then
        tSearch = searchValue.Value
        Set searchValue = Nothing
    Else
        tSearch = searchValue
        searchValue = Null
    End If
    
    '# If searchValue is an array - allow 1D only, it not, error. All arrays from range objects are 2D.
    '  test all arrays and convert 1 by n or n by 1 arrays to n arrays
    If IsArray(tSearch) Then
        searchBase = IIf(LBound(tSearch) = 0, 1, 0)
        
        '# sort out or convert to 1D array
        If Not IsOneDimension(tSearch) Then
            nnRow = UBound(tSearch, 1) + searchBase
            nnCol = UBound(tSearch, 2) + searchBase
            
            '# throw errors for 2D arrays
            If (nnRow > 1) And (nnCol > 1) Then
                MsgBox "No matrix is allowed. It should be either column or row vector as an input argument"
                Exit Function
            Else
                If (nnRow = 1) Or (nnCol = 1) Then
                    tSearch = WorksheetFunction.Transpose(tSearch)
                End If
            End If
        End If
        
        '# searchValue element type: -1 for object, 0 for numeric, otherwise string
        nSearch = WorksheetFunction.Max(nnRow, nnCol)
        nnRow = Null
        nnCol = Null
        searchKeyCount = 0
        For k = 1 To nSearch
            If Not IsNumeric(tSearch(k)) Then
                searchKeyCount = searchKeyCount + 1
            ElseIf IsObject(tSearch(k)) Then
                searchKeyCount = -1
                Exit For
            End If
        Next k
    Else
        nSearch = 1
        searchKeyCount = 0
        If Not IsNumeric(tSearch) Then
            searchKeyCount = searchKeyCount + 1
        ElseIf IsObject(tSearch) Then
            searchKeyCount = -1
        End If
    End If
    
    
    '## target Ref
    '# Input Arguement Type - Range Object? and If not Array, then error
    If TypeName(targetRef) = "Range" Then
        tArr = targetRef.Value
        Set targetRef = Nothing
        
        '# targetRef - array test
        If Not IsArray(tArr) Then
            MsgBox "Out of Scope: Rerence Array is requried", vbExclamation
            PositionInArray = -1
            Exit Function
        Else
            If IsEmpty(tArr) Then
                MsgBox "Reference array is empty"
                PositionInArray = 0
                Exit Function
            Else
                If IsError(UBound(tArr, 2)) Then
                    tDirect = "row"
                    tArr = sMATOOLS.Convert1DTo2DVector(tArr, 1)
                    refBase = IIf(LBound(tArr) = 0, 1, 0)
                    nRow = UBound(tArr, 1) + refBase       ' Measuring The length of the rows of a Matrix
                    nCol = UBound(tArr, 2) + refBase       ' Measuring The length of the size of columns of a Matrix
                Else
                    refBase = IIf(LBound(tArr) = 0, 1, 0)
                    nRow = UBound(tArr, 1) + refBase       ' Measuring The length of the rows of a Matrix
                    nCol = UBound(tArr, 2) + refBase       ' Measuring The length of the size of columns of a Matrix
                    
                    If nRow = 1 Then
                        tDirect = "row"
                    ElseIf nCol = 1 Then
                        tDirect = "column"
                    Else
                        tDirect = "matrix"
                    End If
                End If
            End If
        End If
    ElseIf IsObject(targetRef) Then
        MsgBox "Out of Scope: Reference should be either range or array", vbInformation
        Exit Function
    Else
        '# targetRef - array test
        If Not IsArray(targetRef) Then
            MsgBox "Out of Scope: Rerence Array is requried", vbExclamation
            PositionInArray = -1
            Exit Function
        Else
            If IsEmpty(targetRef) Then
                MsgBox "Reference array is empty"
                PositionInArray = 0
                Exit Function
            Else
                If IsError(UBound(targetRef, 2)) Then
                    tDirect = "row"
                    tArr = sMATOOLS.Convert1DTo2DVector(targetRef, 1)
                    refBase = IIf(LBound(tArr) = 0, 1, 0)
                    nRow = UBound(tArr, 1) + refBase       ' Measuring The length of the rows of a Matrix
                    nCol = UBound(tArr, 2) + refBase       ' Measuring The length of the size of columns of a Matrix
                Else
                    refBase = IIf(LBound(targetRef) = 0, 1, 0)
                    nRow = UBound(targetRef, 1) + refBase       ' Measuring The length of the rows of a Matrix
                    nCol = UBound(targetRef, 2) + refBase       ' Measuring The length of the size of columns of a Matrix
                    
                    If nRow = 1 Then
                        tDirect = "row"
                    ElseIf nCol = 1 Then
                        tDirect = "column"
                    Else
                        tDirect = "matrix"
                    End If
                    
                    ReDim tArr(1 To nRow, 1 To nCol)
                    For i = 1 To nRow
                        For j = 1 To nCol
                            tArr(i, j) = targetRef(i - refBase, j - refBase)
                        Next j
                    Next i
                End If
            End If
            targetRef = Null
        End If
    End If
    
    
    '## Declare the dimmension of the result
    If idxRef = 0 Then
        ReDim tAns(1 To 2 * nSearch, 1 To nRow * nCol)
    Else
        Select Case tDirect
            Case "row"
                ReDim tAns(1 To 2 * nSearch, 1 To nCol)
                idxRef = 1
                ' Column Index
                If IsMissing(resultKey) Or (resultKey = 1) Then
                    searchType = 2
                End If
            Case "column"
                ReDim tAns(1 To 2 * nSearch, 1 To nRow)
                idxRef = 1
                'Row Index
                If IsMissing(resultKey) Or (resultKey = 2) Then
                    resultKey = 1
                End If
            Case "matrix"
                ReDim tAns(1 To 2 * nSearch, 1 To nRow * nCol)
                If IsMissing(resultKey) Then
                    resultKey = 0
                End If
            Case Else
                MsgBox "Reference Array couldn't be scalar.", vbExclamation
                Exit Function
        End Select
    End If
            
    '## Search Indices
    ' Declare for all elements are matched and first row is row index and
    ' second row is column index and will be transposed
    ' Consider idxRef are column index unless Reference is row vector
    ' read column by column with base = 1
    
    '## Largest Count
''    If IsArray(tSearch) Then
''        ReDim tItem(1 To nSerch)
''    End If
    
    '# for all cases
    If idxRef = 0 Then
        For i = 1 To nRow
            For j = 1 To nCol
                ' searchCount - Numeric (0), String or otherwise (-1)
                Select Case searchKeyCount
                    Case 0                          ' Numeric
                        If IsArray(tSearch) Then
                            cCount = 0
                            For k = 1 To nSearch
                                If tSearch(k) = tArr(i, j) Then
                                    cCount = cCount + 1
                                    tAns(1, cCount) = i
                                    tAns(2, cCount) = j
                                End If
                            Next k
                        Else
                            If tSearch = tArr(i, j) Then
                                cCount = cCount + 1
                                tAns(1, cCount) = i
                                tAns(2, cCount) = j
                            End If
                        End If

                    Case Is > 0                    ' String
                        Select Case matchType
                            Case 0
                                If IsArray(tSearch) Then
                                    For k = 1 To nSearch
                                        If tSearch(k) = tArr(i - refBase, j - refBase) Then
                                            cCount = cCount + 1
                                            tAns(1, cCount) = i
                                            tAns(2, cCount) = j
                                        End If
                                    Next k
                                Else
                                    If tSearch = tArr(i - refBase, j - refBase) Then
                                        cCount = cCount + 1
                                        tAns(1, cCount) = i
                                        tAns(2, cCount) = j
                                    End If
                                End If
                            Case -1
                                If IsArray(tSearch) Then
                                    For k = 1 To nSearch
                                        If LCase(tSearch(k)) = LCase(tArr(i - refBase, j - refBase)) Then
                                            cCount = cCount + 1
                                            tAns(1, cCount) = i
                                            tAns(2, cCount) = j
                                        End If
                                    Next k
                                Else
                                    If LCase(tSearch) = LCase(tArr(i - pos, j - pos)) Then
                                        cCount = cCount + 1
                                        tAns(1, cCount) = i
                                        tAns(2, cCount) = j
                                    End If
                                End If
                            Case Is > 1
                                 If IsArray(tSearch) Then
                                    For k = 1 To nSearch
                                        If Left(LCase(tSearch(k)), matchType) = Left(LCase(tArr(i - refBase, j - refBase)), matchType) Then
                                            cCount = cCount + 1
                                            tAns(1, cCount) = i
                                            tAns(2, cCount) = j
                                        End If
                                    Next k
                                Else
                                    If Left(LCase(tSearch), matchType) = Left(LCase(tArr(i - refBase, j - refBase)), matchType) Then
                                        cCount = cCount + 1
                                        tAns(1, cCount) = i
                                        tAns(2, cCount) = j
                                    End If
                                End If
                           Case Else
                                MsgBox "Out of Scope"
                                PositionInArray = -1
                                Exit Function
                               
                        End Select
                    Case Else                     ' Object
                        MsgBox "Out of Scope: Only String (date format) or Numeric Values are valid"
                        PositionInArray = -1
                        Exit Function
                        ''Else
                        ''    MsgBox "Out of Scope"
                        ''    PositionInArray = -1
                        ''    Exit Function
                        ''End If
                    ''End If
                End Select
            Next j
        Next i
        
    ' for a specific column reference
    Else
        For i = 1 To nRow
            For j = 1 To nCol
                ' searchCount
                Select Case searchKeyCount
                    Case 0                      ' Numeric
                        If IsArray(tSearch) Then
                            For k = 1 To nSearch
                                Select Case tDirect
                                    Case "row"
                                        If tSearch(k) = tArr(idxRef, j) Then
                                            cCount = cCount + 1
                                            tAns(1, cCount) = idxRef
                                            tAns(2, cCount) = j
                                        End If
                                    Case "column"
                                        If tSearch(k) = tArr(i, idxRef) Then
                                            cCount = cCount + 1
                                            tAns(1, cCount) = i
                                            tAns(2, cCount) = idxRef
                                        End If
                                    Case "matrix"
                                        If tSearch(k) = tArr(i, idxRef) Then
                                            cCount = cCount + 1
                                            tAns(1, cCount) = i
                                            tAns(2, cCount) = idxRef
                                        End If
                                End Select
                            Next k
                        Else
                            Select Case tDirect
                                Case "row"
                                    If tSearch = tArr(idxRef, j) Then
                                        cCount = cCount + 1
                                        tAns(1, cCount) = idxRef
                                        tAns(2, cCount) = j
                                    End If
                                Case "column"
                                    If tSearch = tArr(i, idxRef) Then
                                        cCount = cCount + 1
                                        tAns(1, cCount) = i
                                        tAns(2, cCount) = idxRef
                                    End If
                                Case "matrix"
                                    If tSearch = tArr(i, idxRef) Then
                                        cCount = cCount + 1
                                        tAns(1, cCount) = i
                                        tAns(2, cCount) = idxRef
                                    End If
                            End Select
                        End If

                    Case Is > 0                     ' String
                    ''Else
                        ''If IsString(searchValue) Then
                            Select Case matchType
                                Case 0
                                    If IsArray(tSearch) Then
                                        For k = 1 To nSearch
                                            Select Case tDirect
                                                Case "row"
                                                    If tSearch(k) = tArr(idxRef - refBase, j - refBase) Then
                                                        cCount = cCount + 1
                                                        tAns(1, cCount) = idxRef
                                                        tAns(2, cCount) = j
                                                    End If
                                                Case "column"
                                                    If tSearch(k) = tArr(i - refBase, idxRef - refBase) Then
                                                        cCount = cCount + 1
                                                        tAns(1, cCount) = i
                                                        tAns(2, cCount) = idxRef
                                                    End If
                                                Case "matrix"
                                                    If tSearch(k) = tArr(i - refBase, idxRef - refBase) Then
                                                        cCount = cCount + 1
                                                        tAns(1, cCount) = i
                                                        tAns(2, cCount) = idxRef
                                                    End If
                                            End Select
                                        Next k
                                    Else
                                        Select Case tDirect
                                            Case "row"
                                                If tSearch = tArr(idxRef - refBase, j - refBase) Then
                                                    cCount = cCount + 1
                                                    tAns(1, cCount) = idxRef
                                                    tAns(2, cCount) = j
                                                End If
                                            Case "column"
                                                If tSearch = tArr(i - refBase, idxRef - refBase) Then
                                                    cCount = cCount + 1
                                                    tAns(1, cCount) = i
                                                    tAns(2, cCount) = idxRef
                                                End If
                                            Case "matrix"
                                                If tSearch = tArr(i - refBase, idxRef - refBase) Then
                                                    cCount = cCount + 1
                                                    tAns(1, cCount) = i
                                                    tAns(2, cCount) = idxRef
                                                End If
                                        End Select
                                    End If
                                Case -1
                                    If IsArray(tSearch) Then
                                        For k = 1 To nSearch
                                            Select Case tDirect
                                                Case "row"
                                                    If LCase(tSearch(k)) = LCase(tArr(idxRef - refBase, j - refBase)) Then
                                                        cCount = cCount + 1
                                                        tAns(1, cCount) = idxRef
                                                        tAns(2, cCount) = j
                                                    End If
                                                Case "column"
                                                    If LCase(tSearch(k)) = LCase(tArr(i - refBase, idxRef - refBase)) Then
                                                        cCount = cCount + 1
                                                        tAns(1, cCount) = i
                                                        tAns(2, cCount) = idxRef
                                                    End If
                                                Case "matrix"
                                                    If LCase(tSearch(k)) = LCase(tArr(i - refBase, idxRef - refBase)) Then
                                                        cCount = cCount + 1
                                                        tAns(1, cCount) = i
                                                        tAns(2, cCount) = idxRef
                                                    End If
                                            End Select
                                        Next k
                                    Else
                                        Select Case tDirect
                                            Case "row"
                                                If LCase(tSearch) = LCase(tArr(idxRef - refBase, j - refBase)) Then
                                                    cCount = cCount + 1
                                                    tAns(1, cCount) = idxRef
                                                    tAns(2, cCount) = j
                                                End If
                                            Case "column"
                                                If LCase(tSearch) = LCase(tArr(i - refBase, idxRef - refBase)) Then
                                                    cCount = cCount + 1
                                                    tAns(1, cCount) = i
                                                    tAns(2, cCount) = idxRef
                                                End If
                                            Case "matrix"
                                                If LCase(tSearch) = LCase(tArr(i - refBase, idxRef - refBase)) Then
                                                    cCount = cCount + 1
                                                    tAns(1, cCount) = i
                                                    tAns(2, cCount) = idxRef
                                                End If
                                        End Select
                                    End If
                                Case Is > 1
                                     If IsArray(tSearch) Then
                                            Select Case tDirect
                                                Case "row"
                                                    If Left(LCase(tSearch(k)), matchType) = Left(LCase(tArr(idxRef - refBase, j - refBase)), matchType) Then
                                                        cCount = cCount + 1
                                                        tAns(1, cCount) = idxRef
                                                        tAns(2, cCount) = j
                                                    End If
                                                Case "column"
                                                    If Left(LCase(tSearch(k)), matchType) = Left(LCase(tArr(i - refBase, idxRef - refBase)), matchType) Then
                                                        cCount = cCount + 1
                                                        tAns(1, cCount) = i
                                                        tAns(2, cCount) = idxRef
                                                    End If
                                                Case "matrix"
                                                    If Left(LCase(tSearch(k)), matchType) = Left(LCase(tArr(i - refBase, idxRef - refBase)), matchType) Then
                                                        cCount = cCount + 1
                                                        tAns(1, cCount) = i
                                                        tAns(2, cCount) = idxRef
                                                    End If
                                            End Select
                                    Else
                                        Select Case tDirect
                                            Case "row"
                                                If Left(LCase(tSearch), matchType) = Left(LCase(tArr(idxRef - refBase, j - refBase)), matchType) Then
                                                    cCount = cCount + 1
                                                    tAns(1, cCount) = idxRef
                                                    tAns(2, cCount) = j
                                                End If
                                            Case "column"
                                                If Left(LCase(tSearch), matchType) = Left(LCase(tArr(i - refBase, idxRef - refBase)), matchType) Then
                                                    cCount = cCount + 1
                                                    tAns(1, cCount) = i
                                                    tAns(2, cCount) = idxRef
                                                End If
                                            Case "matrix"
                                                If Left(LCase(tSearch), matchType) = Left(LCase(targetRef(i - refBase, idxRef - refBase)), matchType) Then
                                                    cCount = cCount + 1
                                                    tAns(1, cCount) = i
                                                    tAns(2, cCount) = idxRef
                                                End If
                                        End Select
                                    End If
                               Case Else
                                    MsgBox "Out of Scope"
                                    PositionInArray = -1
                                    Exit Function
                                   
                            End Select
                    Case Else                     ' Object
                        MsgBox "Out of Scope: Only String (date format) or Numeric Values are valid"
                        PositionInArray = -1
                        Exit Function
                End Select
            Next j
        Next i
    End If
                
                
'''                    For i = 1 To nRow
'''                         If IsNumeric(searchValue) Then
'''                             If searchValue = targetRef(i - pos, idxRef - pos) Then
'''                                 cCount = cCount + 1
'''                                 tAns(1, cCount) = i                     ' row
'''                                 tAns(2, cCount) = idxRef                ' column
'''                             End If
'''                         Else
'''                             If IsString(searchValue) Then
'''                                 Select Case matchType
'''                                     Case 0
'''                                         If searchValue = targetRef(i - pos, idxRef - pos) Then
'''                                             cCount = cCount + 1
'''                                             tAns(1, cCount) = i
'''                                             tAns(2, cCount) = idxRef
'''                                         End If
'''                                     Case -1
'''                                         If LCase(searchValue) = LCase(targetRef(i - pos, idxRef - pos)) Then
'''                                             cCount = cCount + 1
'''                                             tAns(1, cCount) = i
'''                                             tAns(2, cCount) = idxRef
'''                                         End If
'''                                     Case 1 To nRow
'''                                          If Left(LCase(searchValue), matchType) = Left(LCase(targetRef(i - pos, idxRef - pos)), matchType) Then
'''                                             cCount = cCount + 1
'''                                             tAns(1, cCount) = i
'''                                             tAns(2, cCount) = idxRef
'''                                         End If
'''                                    Case Is > nRow
'''                                         MsgBox "Out of Scope"
'''                                         PositionInArray = -1
'''                                         Exit Function
'''                                 End Select
'''                             Else
'''                                 MsgBox "Out of Scope"
'''                                 PositionInArray = -1
'''                                 Exit Function
'''                             End If
'''                         End If
'''                     Next i
'''                End If
            
'''                If IsArray(tSearch) Then
'''                    For k = 1 To WorksheetFunction.Max(nnRow, nnCol)
'''                         If tSearch(k) = tArr(i, j) Then
'''                            cCount = cCount + 1
'''                            tAns(1, cCount) = i
'''                            tAns(2, cCount) = j
'''                        End If
'''                    Next k
'''                Else
'''                    If tSearch = tArr(i, j) Then
'''                        cCount = cCount + 1
'''                        tAns(1, cCount) = i
'''                        tAns(2, cCount) = j
'''                    End If
'''                End If
''            Next j
''
''        ' specific columns
''        Else
'''            If IsArray(tSearch) Then
'''                For k = 1 To WorksheetFunction.Max(nnRow, nnCol)
'''                    If tSearch(k) = tArr(i, idxRef) Then
'''                        cCount = cCount + 1
'''                        tAns(1, cCount) = i
'''                        tAns(2, cCount) = idxRef
'''                    End If
'''                Next k
'''            Else
'''
'''                If searchValue = tArr(i, idxRef) Then
'''                    cCount = cCount + 1
'''                    tAns(1, cCount) = i
'''                    tAns(2, cCount) = idxRef
'''                End If
'''            End If
'''        End If
'''    Next i
'''
'''    ReDim Preserve tAns(2, cCount)
'''
'''    ' VBA Array Reference
'''    Else
'''        If IsEmpty(targetRef) Then
'''            MsgBox "Reference array is empty"
'''            PositionInArray = 0
'''            Exit Function
'''        Else
'''            ' 1D Reference Array
'''            If sMATOOLS.IsOneDimension(targetRef) Then
'''                pos = IIf(LBound(targetRef) = 0, 1, 0)
'''                nCol = UBound(targetRef) + pos                          ' Measuring The length of the rows of a Matrix
'''                idxRef = 1
'''                ' Declare for all elements are matched and first row is row index and second row is column index and will be transposed
'''
'''                ReDim tAns(1 To 2, 1 To nCol)
'''
'''                ' read column by column with base = 1
'''                For j = 1 To nCol
'''                    If IsArray(tSearch) Then
'''                        For k = 1 To WorksheetFunction.Max(nnRow, nnCol)
'''                            If tSearch(k) = targetRef(j) Then
'''                                cCount = cCount + 1
'''                                tAns(1, cCount) = j
'''                                tAns(2, cCount) = idxRef
'''                            End If
'''                        Next k
'''                    Else
'''                        If tSearch = targetRef(j) Then
'''                            cCount = cCount + 1
'''                            tAns(1, cCount) = j
'''                            tAns(2, cCount) = idxRef
'''                        End If
'''                    End If
'''                Next j
'''            ' 2D Reference Array
'''            Else
'''                ' Sort out the empty Array
'''                If IsEmpty(targetRef) Then
'''                    MsgBox "Reference array is required"
'''                    PositionInArray = 0
'''                    Exit Function
'''                End If
'''
'''                ' Base Test
'''                pos = IIf(LBound(targetRef) = 0, 1, 0)
'''                nRow = UBound(targetRef, 1) + pos               ' Measuring The length of the rows of a Matrix
'''                nCol = UBound(targetRef, 2) + pos               ' Measuring The length of the size of columns of a Matrix
'''
'''                ' Declare for all elements are matched and first row is row index and second row is column index and will be transposed
'''                If idxRef = 0 Then
'''                    ReDim tAns(1 To 2, 1 To nRow * nCol)
'''                Else
'''                    If nCol > 1 Then
'''                        If nRow = 1 Then
'''                            ReDim tAns(1 To 2, 1 To nCol)
'''                        Else
'''                            ReDim tAns(1 To 2, nRow * nCol)
'''                        End If
'''                    Else
'''                        If nRow = 1 Then
'''                            MsgBox "Out of Scope: Reference Array should not be a scalar"
'''                            Exit Function
'''                        Else
'''                            ReDim tAns(1 To 2, 1 To nRow)
'''                        End If
'''                    End If
'''                End If
                
'''                ' read column by column with base = 1
'''                If idxRef = 0 Then
'''                    For i = 1 To nRow
'''                        For j = 1 To nCol
'''                            If IsNumeric(tSearch) Then
'''                                If IsArray(tSearch) Then
'''                                    For k = 1 To WorksheetFunction.Max(nnRow, nnCol)
'''                                        If tSearch(k) = targetRef(j) Then
'''                                            cCount = cCount + 1
'''                                            tAns(1, cCount) = j
'''                                            tAns(2, cCount) = idxRef
'''                                        End If
'''                                    Next k
'''                                Else
'''                                    If tSearch = targetRef(j) Then
'''                                        cCount = cCount + 1
'''                                        tAns(1, cCount) = j
'''                                        tAns(2, cCount) = idxRef
'''                                    End If
'''                                End If
'''
'''
'''
'''                                If searchValue = targetRef(i - pos, j - pos) Then
'''                                    cCount = cCount + 1
'''                                    tAns(1, cCount) = i
'''                                    tAns(2, cCount) = j
'''                                End If
'''                            Else
'''                                If IsString(searchValue) Then
'''                                    Select Case matchType
'''                                        Case 0
'''                                            If searchValue = targetRef(i - pos, j - pos) Then
'''                                                cCount = cCount + 1
'''                                                tAns(1, cCount) = i
'''                                                tAns(2, cCount) = j
'''                                            End If
'''                                        Case -1
'''                                            If LCase(searchValue) = LCase(targetRef(i - pos, j - pos)) Then
'''                                                cCount = cCount + 1
'''                                                tAns(1, cCount) = i
'''                                                tAns(2, cCount) = j
'''                                            End If
'''                                        Case Is > 1
'''                                             If Left(LCase(searchValue), matchType) = Left(LCase(targetRef(i - pos, j - pos)), matchType) Then
'''                                                cCount = cCount + 1
'''                                                tAns(1, cCount) = i
'''                                                tAns(2, cCount) = j
'''                                            End If
'''                                       Case Else
'''                                            MsgBox "Out of Scope"
'''                                            PositionInArray = -1
'''                                            Exit Function
'''
'''                                    End Select
'''                                Else
'''                                    MsgBox "Out of Scope"
'''                                    PositionInArray = -1
'''                                    Exit Function
'''                                End If
'''                            End If
'''                        Next j
'''                    Next i
'''
'''                ' for a specific column reference
'''                Else
'''                    For i = 1 To nRow
'''                         If IsNumeric(searchValue) Then
'''                             If searchValue = targetRef(i - pos, idxRef - pos) Then
'''                                 cCount = cCount + 1
'''                                 tAns(1, cCount) = i                     ' row
'''                                 tAns(2, cCount) = idxRef                ' column
'''                             End If
'''                         Else
'''                             If IsString(searchValue) Then
'''                                 Select Case matchType
'''                                     Case 0
'''                                         If searchValue = targetRef(i - pos, idxRef - pos) Then
'''                                             cCount = cCount + 1
'''                                             tAns(1, cCount) = i
'''                                             tAns(2, cCount) = idxRef
'''                                         End If
'''                                     Case -1
'''                                         If LCase(searchValue) = LCase(targetRef(i - pos, idxRef - pos)) Then
'''                                             cCount = cCount + 1
'''                                             tAns(1, cCount) = i
'''                                             tAns(2, cCount) = idxRef
'''                                         End If
'''                                     Case 1 To nRow
'''                                          If Left(LCase(searchValue), matchType) = Left(LCase(targetRef(i - pos, idxRef - pos)), matchType) Then
'''                                             cCount = cCount + 1
'''                                             tAns(1, cCount) = i
'''                                             tAns(2, cCount) = idxRef
'''                                         End If
'''                                    Case Is > nRow
'''                                         MsgBox "Out of Scope"
'''                                         PositionInArray = -1
'''                                         Exit Function
'''                                 End Select
'''                             Else
'''                                 MsgBox "Out of Scope"
'''                                 PositionInArray = -1
'''                                 Exit Function
'''                             End If
'''                         End If
'''                     Next i
'''                End If
'''            End If
'''        End If
'''    End If
                
    ReDim Preserve tAns(1 To 2, 1 To cCount)
    
    Select Case searchType
        Case "all", 0
            Select Case resultKey
                Case 0
                    PositionInArray = tAns                                  'WorksheetFunction.Transpose(tAns)
                Case 1
                    PositionInArray = WorksheetFunction.Index(tAns, 1, 0)   'WorksheetFunction.Transpose(WorksheetFunction.Index(tAns, 0, 1))
                Case 2
                    PositionInArray = WorksheetFunction.Index(tAns, 2, 0)   'WorksheetFunction.Transpose(WorksheetFunction.Index(tAns, 1, 0))
'                Case Else
'                    If tDirect = "row" Then
'                        PositionInArray = WorksheetFunction.Index(tAns, 0, 1)
'                    Else
'                        PositionInArray = WorksheetFunction.Index(tAns, 1, 0)
'                    End If
            End Select
                
        Case "first", "begin", 1
            Select Case resultKey
                Case 0
                    PositionInArray = WorksheetFunction.Transpose(Array(tAns(1, 1), tAns(2, 1)))
                Case 1
                    PositionInArray = tAns(1, 1)
                Case 2
                    PositionInArray = tAns(2, 1)
'                Case Else
'                    If tDirect = "row" Then
'                        PositionInArray = tAns(1, 1)
'                    Else
'                        PositionInArray = tAns(2, 1)
'                    End If
            End Select
            
        Case "last", "end"
            Select Case resultKey
                Case 0
                    PositionInArray = WorksheetFunction.Transepose(Array(tAns(1, cCount), tAns(2, cCount)))
                Case 1
                    PositionInArray = tAns(1, cCount)
                Case 2
                    PositionInArray = tAns(2, cCount)
'                Case Else
'                    If tDirect = "row" Then
'                        PositionInArray = tAns(1, cCount)
'                    Else
'                        PositionInArray = tAns(2, cCount)
'                    End If
            End Select
            
        Case Else
            If Not IsString(searchType) Then
                PositionInArray = tAns(searchType)
            Else
                MsgBox "Not right format"
                Exit Function
            End If
    End Select
End Function


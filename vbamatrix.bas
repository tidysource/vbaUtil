Attribute VB_Name = "matrix_"
Option Explicit

'Get a matrix of named rows and columns
'---------------------------------------
Function namedMatrix( _
                        matrix As Variant, _
                        Optional columnNames As Variant, _
                        Optional rowNames As Variant _
                        )
    'Matrix must have at least one row and one column
    If length(matrix) < 1 _
    Or length(matrix(0)) < 1 Then
        Err.Raise 13
    End If
    
    'Default columnNames
    Dim i As Long
    Dim startI As Integer
    startI = 0
    If IsMissing(columnNames) Then
        Dim firstRow As Variant
        firstRow = matrix(0)
        ReDim columnNames(LBound(firstRow) To UBound(firstRow)) As String
        For i = LBound(firstRow) To UBound(firstRow)
            columnNames(i) = firstRow(i)
        Next i
        startI = 1
    End If
    
    'Build result
    Dim j As Long
    Dim k As Long
    Dim row As Collection
    Dim rowArr As Variant

    Dim result As New Collection
    If IsMissing(rowNames) Then
        Dim rowVals As Variant
        For i = LBound(matrix) To UBound(matrix)
            rowArr = matrix(i + startI)
            
            'Build row collection
            Set row = New Collection
            rowVals = slice(rowArr, LBound(rowArr) + 1)
            j = LBound(rowVals)
            For k = LBound(columnNames) To UBound(columnNames)
                row.Add _
                        Key:=columnNames(k), _
                        Item:=rowVals(j)
                j = j + 1
            Next k
            
            'Add row to result collection
            result.Add _
                        Key:=rowArr(LBound(rowArr)), _
                        Item:=row
        Next i
    ElseIf IsArray(rowNames) = True Then
        j = LBound(matrix)

        For i = LBound(rowNames) To UBound(rowNames)
            rowArr = matrix(j + startI)
            j = j + 1
            
            'Build row
            Set row = New Collection
            Dim l As Long
            l = LBound(rowArr)
            For k = LBound(columnNames) To UBound(columnNames)
                row.Add _
                       Key:=columnNames(k), _
                       Item:=rowArr(l)
                l = l + 1
            Next k
          
            'Add row to result collection
            result.Add _
                        Key:=rowNames(i), _
                        Item:=row
        Next i
    Else
        Err.Raise 13
    End If
    
    Set namedMatrix = result
End Function

'Get a matrix of named columns
'-----------------------------
Function namedColumns( _
                        matrix As Variant, _
                        Optional names As Variant _
                        ) As Variant
    'Matrix must have at least one row and one column
    If length(matrix) < 1 _
    Or length(matrix(0)) < 1 Then
        Err.Raise 13
    End If
    
    'Default names array
    Dim i As Long
    Dim startI As Integer
    startI = 0
    If IsMissing(names) Then
        Dim firstRow As Variant
        firstRow = matrix(0)
        ReDim names(LBound(firstRow) To UBound(firstRow)) As String
        For i = LBound(firstRow) To UBound(firstRow)
            names(i) = firstRow(i)
        Next i
        startI = 1
    End If
        
    'Build result
    Dim j As Long
    Dim k As Long
    
    Dim result As Variant
    ReDim result(LBound(matrix) + startI To UBound(matrix)) As Variant
    
    For i = LBound(matrix) + startI To UBound(matrix)
        Dim row As Collection
        Set row = New Collection
        k = LBound(matrix(i))
        For j = LBound(names) To UBound(names)
            row.Add _
                Key:=names(j), _
                Item:=matrix(i)(k)
            k = k + 1
        Next j
        Set result(i) = row
    Next i
    
    namedColumns = result
End Function


'Get a matrix of named rows
'--------------------------
Function namedRows( _
                    matrix As Variant, _
                    Optional names As Variant _
                    ) As Variant
    Dim result As New Collection
    Dim i As Long
    Dim j As Long
    Dim row As Variant
    
    If IsMissing(names) Then
        For i = LBound(matrix) To UBound(matrix)
            row = matrix(i)
            
            If emptyArr(row) = False Then 'skip empty rows
                j = LBound(row)
                result.Add _
                        Key:=row(j), _
                        Item:=slice(row, 1)
            End If
        Next i
    ElseIf IsArray(names) = True Then
        j = LBound(matrix)
        For i = LBound(names) To UBound(names)
            row = matrix(j)
            result.Add _
                        Key:=names(i), _
                        Item:=row
                        
            j = j + 1
        Next i
    Else
        'Optional names argument must be an array
        Err.Raise 13
    End If
    
    Set namedRows = result
End Function


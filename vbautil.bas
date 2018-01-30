Attribute VB_Name = "array_"
Option Explicit

'Wrap string in quotes
'---------------------
Function quote(str As String)
    quote = """" & str & """"
End Function

'True/false if array is empty
'----------------------------
Function emptyArr(arr As Variant) As Boolean
    If LBound(arr) > UBound(arr) Then
        emptyArr = True
    Else
        emptyArr = False
    End If
End Function

'Print out array values in immediate window
'Note: array may contain only simple values (no objects or other arrays)
Function printArr(arr As Variant)
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        Debug.Print arr(i)
    Next i
End Function

'Get length of an array
'----------------------
Function countArr(arr As Variant) As Long
    countArr = UBound(arr) - LBound(arr) + 1
End Function

'Change base of an array
'-----------------------
Function rebaseArr(arr As Variant, _
                    baseIndex As Long _
                    ) As Variant
    Dim result As Variant
    
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        result = arr(i)
    Next i
    
    rebaseArr = result
End Function

'Get lbound of array/collection/string
'-------------------------------------
Function baseIndex(param) As Long
    If IsArray(param) = True Then
        baseIndex = LBound(param)
    ElseIf TypeName(param) = "Collection" Then
        baseIndex = 1
    ElseIf TypeName(param) = "String" Then
        baseIndex = 1
    Else
        Err.Raise 13
    End If
End Function

'Get ubound of array/collection/string
'-------------------------------------
Function lastIndex(param) As Long
    If IsArray(param) = True Then
        lastIndex = UBound(param)
    ElseIf TypeName(param) = "Collection" Then
        lastIndex = param.Count
    ElseIf TypeName(param) = "String" Then
        lastIndex = Len(param)
    Else
        Err.Raise 13
    End If
End Function

'Get length of array/collection/string
'-------------------------------------
Function length(param) As Long
    If IsArray(param) = True Then
        length = countArr(param)
    Else
        Dim paramType As String
        paramType = TypeName(param)
                
        If paramType = "Collection" Then
            length = param.Count
        ElseIf paramType = "String" Then
            length = Len(param)
        Else
            Err.Raise 13
        End If
    End If
End Function


'Return a subset of array/collection/string
'------------------------------------------
'endIndex is not included
'base 0 is assumed for array, collection and string indexes
'
'IMPORTANT:
'does not return a copy of collection,
'the input collection is altered to preserve keys
Function slice( _
                param As Variant, _
                Optional startIndex As Long = 0, _
                Optional endIndex As Variant _
                ) As Variant
    'Set startIndex
    If startIndex < 0 Then
        startIndex = length(param) + startIndex
    End If
    
    'Set endIndex
    If IsMissing(endIndex) Then
        endIndex = length(param)
    ElseIf endIndex < 0 Then
        endIndex = length(param) + endIndex
    End If
    
    'Length of output
    Dim l As Long
    l = endIndex - startIndex
    
    'endIndex is excluded
    endIndex = endIndex - 1
    
    'Get result
    Dim result As Variant
    Dim i As Long
    
    If IsArray(param) Then
        'Rebase to array bounds
        startIndex = startIndex + LBound(param)
        endIndex = endIndex + LBound(param)
        Dim j As Long
        j = LBound(param) 'preserve input base
        ReDim result(j To j + l - 1) As Variant
        For i = startIndex To endIndex
            result(j) = param(i)
            j = j + 1
        Next i
        slice = result
    ElseIf TypeName(param) = "String" Then
        'Rebase to 1
        startIndex = startIndex + 1
        endIndex = endIndex + 1
        If startIndex <= endIndex Then
            slice = Mid(param, startIndex, 1)
        Else
            slice = ""
        End If
    ElseIf TypeName(param) = "Collection" Then
        'Rebase to 1
        startIndex = startIndex + 1
        endIndex = endIndex + 1
        For i = param.Count To 1 Step -1
            If i < startIndex Or i > endIndex Then
                param.Remove i
            End If
        Next i
        slice = param
    Else
        Err.Raise 13
    End If
End Function

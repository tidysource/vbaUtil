Attribute VB_Name = "array_"
Option Explicit

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
        endIndex = length(param) - 1 + endIndex
    End If
    
    Dim result As Variant
    Dim i As Long
    
    If IsArray(param) = True Then
        If startIndex <= endIndex Then
            'Adjust to array bounds
            startIndex = startIndex + LBound(param)
            endIndex = endIndex + LBound(param)
              
            Dim j As Long
            j = LBound(param)
            ReDim result(j To endIndex - 1 - startIndex) As Variant
            For i = startIndex To endIndex - 1
                result(j) = param(i)
                j = j + 1
            Next i
        Else
            'Return an empty array
            ReDim result(0 To -1)
        End If
        'Return subset
        slice = result
    ElseIf TypeName(param) = "Collection" Then
        For i = param.Count To 1 Step -1
            If i < startIndex Or i > endIndex Then
                param.Remove i
            End If
        Next i
        
        slice = param
    ElseIf TypeName(param) = "String" Then
        result = ""
        
        If startIndex <= endIndex Then
            result = Mid(param, startIndex + 1, endIndex)
        End If
        
        'Return subset
        slice = result
    Else
        Err.Raise 13
    End If
End Function

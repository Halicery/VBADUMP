Attribute VB_Name = "BUG_VBA64ParamArray"
' 64-bit VBA ParamArray LongLong ByRef BUG
' ----------------------------------------
' (c) 2024 A. Tarpai https://github.com/Halicery/VBADUMP/
'
'
' There is a VBA64 bug or missing feature
' Cannot write/modify LongLong variable passed in ParamArray
' Can modify all other variable data types
' Can modify all other LongLong expressions, literals and function return values

' By narrowing down the problem lies in VT_BYREF-Variant in ParamArray:
'
' 0240  vbInteger|VT_BYREF     OK
' 0340  vbLong|VT_BYREF        OK
' 1440  vbLongLong|VT_BYREF    ERROR
'
' See also:
' https://vba-internal-investigations.blogspot.com/2024/03/vba-paramarray-and-built-in-array.html
'

Option Explicit

#If Win64 Then

    'Private Sub pLL(LL() As LongLong)
    '    Debug.Print LL(0), VarType(LL(0)), TypeName(LL(0))
    '    LL(0) = -1 ' modify: OK
    'End Sub
    
    Private Sub pa(ParamArray a())
        Debug.Print a(0), VarType(a(0)), TypeName(a(0))
        a(0) = -1 ' modify: ERROR
    End Sub
    
    Private Sub passParamArray()
        Dim i As Integer
        Dim L As Long
        Dim LL As LongLong
        
        pa i   ' OK
        pa L   ' OK
        'pa LL  ' <-- "Variable uses an Automation type not supported in Visual Basic (Error 458)"
        
        ' literals, expressions, return values: all OK (not VT_BYREF Variant)
        pa 100^             ' LongLong literal
        pa CLngLng(8)       ' returns LongLong
        pa VarPtr(LL)       ' returns LongPtr = LongLong
        pa f_ret_longlong   ' returns LongLong
        
        Debug.Print "I", i
        Debug.Print "L", L
        Debug.Print "LL", LL
    End Sub
    
    Private Function f_ret_longlong() As LongLong
        f_ret_longlong = 100
    End Function

#End If


Attribute VB_Name = "BUG_ArrayAssignment"
' VBA Array Assignment BUG
' ========================
'
' This is a feature or BUG and can be a serious ERROR.
' Affects String- and Variant Arrays returned from VBA Module UDF functions followed by array-assignment.
' Corrupt BSTR buffer and corrupt Heap, can lead to Host App CRASH.
'
' Blog:
' https://vba-internal-investigations.blogspot.com/2024/03/vba-array-assignment-bug.html

' MIT License
'
' Copyright (c) 2023 A. Tarpai https://github.com/Halicery/VBADUMP/
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Test-run: String array
'
' Run this first - safe.
' It is NOT supposed to print the same BSTR address - StrPtr()
' and NOT supposed to change two strings in two arrays
'
' prints:
' arr(0):       !ot me
' arr2(0):      !ot me

Private Sub array_assignment_bug_string_array_test()
    Dim arr() As String, arr2() As String
    arr = return_arr_string
    arr2 = arr
    Debug.Print Hex(StrPtr(arr(0)))
    Debug.Print Hex(StrPtr(arr2(0)))
    Mid(arr(0), 1, 1) = "!"  ' change char
    Debug.Print "arr(0): ", arr(0)
    Debug.Print "arr2(0): ", arr2(0)
End Sub


Private Function return_arr_string() As String()
    ReDim return_arr_string(2) 'As String
    fill_arr_string return_arr_string
End Function

' use a worker Sub to go around VBA's array-indexing syntax problem
Private Sub fill_arr_string(arr() As String)
    arr(0) = "not me"
End Sub



'''''''' Might crash/corrupt BSTR, especially in F8-step Debug
'
Private Sub array_assignment_bug_string_array_crashtest()
    Dim arr() As String, arr2() As String
    arr = return_arr_string
    arr2 = arr
    
    ' do sth involving BSTR release/alloc
    ' Unpredictable, sometimes changes sometimes not. Mainly in Debug
    arr(0) = vbNullString ' release BSTR <-- Set breakpoint here
    arr(0) = "only this" ' <-- Set breakpoint here: NOT SUPPOSED TO CHANGE
    Debug.Print "arr(0): ", arr(0)
    Debug.Print "arr2(0): ", arr2(0)
    'Debug.Print Hex(VarPtr(arr(0)))
    'Debug.Print Hex(VarPtr(arr2(0)))
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Test-run with Variant Array
'
' Here we store a String and another Variant Array containing String in the Variant array:
' Then examine the strings: all the same behaviour
' Ergo: another simple memcopy of elements

Private Sub array_assignment_bug_variant_array_strings()
    Dim arr(), arr2()
    arr = return_arr_variant
    arr2 = arr
    
    ' String in Variant array ("buggy")
    Debug.Print Hex(StrPtr(arr(0)))
    Debug.Print Hex(StrPtr(arr2(0))) ' NOT supposed to be the same address
    Mid(arr(0), 1, 1) = "!"
    Debug.Print "arr(0): ", arr(0)
    Debug.Print "arr2(0): ", arr2(0) ' this was NOT supposed to change
    
    ' in contained array
    'Debug.Print Hex$(StrPtr(arr(1)(0)))  ' NB! a temp BSTR is created and passed to StrPtr()
    'Debug.Print Hex$(StrPtr2(arr2(1)(0))) ' Use StrPtr2() in vbadump.bas
    Mid(arr(1)(0), 1, 1) = "!"
    Debug.Print "arr(1)(0): ", arr(1)(0)
    Debug.Print "arr2(1)(0): ", arr(1)(0) ' this was NOT supposed to change
End Sub

Private Function return_arr_variant() As Variant()
    ReDim return_arr_variant(2)
    fill_arr_variant return_arr_variant
End Function

Private Sub fill_arr_variant(v())  ' fill values: a String and an Array
    v(0) = "buggy"
    v(1) = Array("inside", 5)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' THIS USUALLY CRASHES EXCEL
' Especially in Debug mode and Locals Window alive (sometimes App defined error appears then crash)
'
' Erasing one of the Dynamic Array: corrupt Heap for the other SAFEARRAY pointer and APP CRASH

Private Sub array_assignment_bug_variant_array_crash()
    Dim varr(), varr2()
    varr = return_arr_variant
    varr2 = varr
    
    Erase varr2(1) ' DO NOT EXECUTE
    ' CORRUPT HEAP. CERTIANLY FATAL
  
End Sub


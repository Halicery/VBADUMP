Attribute VB_Name = "vbadump"
' MIT License
'
' Copyright (c) 2023 Attila Tarpai https://github.com/Halicery/VBADUMP/
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
Option Base 0

Public Declare PtrSafe Sub memcpy Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDestination As Any, ByVal pSource As Any, ByVal length As Long)

Private Type SAFEARRAYBOUND
  cElements As Long
  lLbound As Long
End Type

Private Type SAFEARRAY
  cDims As Integer
  fFeatures As Integer
  cbElements As Long
  cLocks As Long
  pvData As LongPtr
  'SAFEARRAYBOUND rgsabound[1];
End Type

Private Enum ADVFEATUREFLAGS  ' fFeatures 16 bits
    FADF_AUTO = &H1
    FADF_STATIC = &H2
    FADF_EMBEDDED = &H4
    FADF_FIXEDSIZE = &H10
    FADF_RECORD = &H20       ' IRecordInfo at -4/8
    FADF_HAVEIID = &H40      ' GUID at -16
    FADF_HAVEVARTYPE = &H80  ' VbVarType at -4
    FADF_BSTR = &H100
    FADF_UNKNOWN = &H200
    FADF_DISPATCH = &H400
    FADF_VARIANT = &H800
End Enum

'Private Type STRUCTVARIANT
'  VTYPE As Integer
'  wReserved1 As Integer
'  wReserved2 As Integer
'  wReserved3 As Integer
'  data As LongPtr
'  pRecInfo As LongPtr
'End Type

Private Enum VARENUMFLAGS
  VT_VECTOR = &H1000
  VT_ARRAY = &H2000
  VT_BYREF = &H4000
  VT_RESERVED = &H8000
  VT_TYPEMASK = &HFFF
End Enum

'Private Type STRUCTDECIMAL
'  wReserved As Integer
'  scale As Byte
'  sign As Byte
'  Hi32 As Long
'  Lo32 As Long
'  Mid32 As Long
'End Type

Private Const FACILITY_MASK = &H7FF0000  '  11 bits. In vbError (HRESULT/SCODE)
Private Const strADVFEATUREFLAGS = "FADF_AUTO,FADF_STATIC,FADF_EMBEDDED,unk_0008,FADF_FIXEDSIZE,FADF_RECORD,FADF_HAVEIID,FADF_HAVEVARTYPE,FADF_BSTR,FADF_UNKNOWN,FADF_DISPATCH,FADF_VARIANT,unk_1000,unk_2000,unk_4000,unk_8000"
Private Const strVARENUMFLAGS = "VT_VECTOR,VT_ARRAY,VT_BYREF,VT_RESERVED"  ' 12-15
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function flags_string(ByVal flag As Long, enumstr As String, ByVal first_bit As Long, ByVal last_bit As Long) As String
    Dim sf() As String
    sf = Split(enumstr, ",")
    Dim i As Long
    For i = last_bit To first_bit Step -1
        If (2 ^ i And flag) Then
            If vbNullString = flags_string Then
                flags_string = sf(i - first_bit)
            Else
                flags_string = flags_string & "|" & sf(i - first_bit)
            End If
        End If
    Next i
    'If vbNullString = flags_string Then flags_string = "NONE"
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Hex writers with leading zeroes

Function hexi(ByVal v As LongPtr, ByVal n_digit As Long) As String
    hexi = Hex$(v)
    If Len(hexi) < n_digit Then
        hexi = String$(n_digit - Len(hexi), "0") & hexi
    ElseIf Len(hexi) > n_digit Then ' neg
        hexi = Right$(hexi, n_digit)
    End If
End Function

' Writes N hex bytes into String from Byte Array
Private Function bbuffdump(bbuff() As Byte, ByVal firstIdx As Long, ByVal n_bytes As Long) As String  ' print in memory order
    Debug.Assert n_bytes > 0  'Then Exit Function
    Dim s As String, i As Long, b As Byte
    s = String$(n_bytes * 2, "0")
    For i = 0 To n_bytes - 1
        b = bbuff(i + firstIdx)
        If b > 0 Then Mid$(s, i * 2 + 1 - (b < 16)) = Hex$(b) ' True=-1
    Next i
    bbuffdump = s
End Function

' From address in memory. Storage order and possible sectioning
Private Function dumpaddr(ByVal addr As LongPtr, ByVal n_bytes As Long, ParamArray part_sizes()) As String '
    Dim bbuff() As Byte
    read_mem addr, n_bytes, bbuff  ' copy n bytes
    Dim v, st As Long
    For Each v In part_sizes
        dumpaddr = dumpaddr & bbuffdump(bbuff, st, v) & "-"
        st = st + v
    Next v
    dumpaddr = dumpaddr & bbuffdump(bbuff, st, n_bytes - st) ' append rest
End Function

Private Sub read_mem(ByVal addr As LongPtr, ByVal n_bytes As Long, ByRef bbuff() As Byte) ' -> into Byte Array
    ' If n_bytes < 0 Or n_bytes > 255 Then n_bytes = 16 ' sanity check?
    ReDim bbuff(n_bytes - 1)
    memcpy VarPtr(bbuff(0)), addr, n_bytes
End Sub

' Pointers
Function hexi_ptr(ByVal addr As LongPtr) As String
    hexi_ptr = hexi(addr, LenB(addr) * 2)
End Function

Function read_ptr(ByVal addr As LongPtr) As LongPtr
    memcpy VarPtr(read_ptr), addr, LenB(addr)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Variant helpers

Private Function IsBYREF(v) As Long
    Dim VTYPE As Integer
    memcpy VarPtr(VTYPE), VarPtr(v), LenB(VTYPE)  ' peek VarType
    IsBYREF = (VTYPE And VT_BYREF)    'If (VTYPE And VT_BYREF) Then IsBYREF = True
End Function

' LenB2()
Private Function sizeof(vt As VbVarType) As Long       ' LenB(v) is LenB(CStr(v)) unfortunately
    Select Case vt
        #If Win64 Then
            Case vbLongLong
                sizeof = 8
        #End If
        Case vbString, vbObject, Is > 8192 ' pointer types
            Dim p As LongPtr
            sizeof = LenB(p)
        Case vbByte
            sizeof = 1
        Case vbInteger, vbBoolean
            sizeof = 2
        Case vbLong, vbSingle, vbError
            sizeof = 4
        Case vbDouble, vbDate, vbCurrency
            sizeof = 8
        Case vbDecimal  ' this is special
            sizeof = 16
        Case Else ' vbEmpty, vbNull, others..
            sizeof = 0
    End Select
End Function

' Modified StrPtr(v) and VarPtr(v)
' Calling StrPtr() for String or Variant/String elements in array in Variant makes a temp BSTR
' VarPtr(v(0)) is also temp-address for Strings (in Variant of array of String)
' -1 means not a String

Function StrPtr2(v) As LongPtr   ' Has to be Variant: Variant/String type mismatch for String
    If vbString = VarType(v) Then
        If IsBYREF(v) Then
            StrPtr2 = StrPtr(v)
        Else
            StrPtr2 = read_ptr(VarPtr(v) + 8) ' extract from Variant? as is
        End If
    Else
        StrPtr2 = -1
    End If
End Function

Function VarPtr2(v) As LongPtr
        If IsBYREF(v) Then
            VarPtr2 = read_ptr(VarPtr(v) + 8) ' extract from TEMP VT_BYREF
        Else
            VarPtr2 = VarPtr(v)
        End If
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' We can pass any variable/expression As ByRef Variant
'
' Variant can be:
' - a Variant
' - VT_BYREF-Variant --> pointer to non-Variant
'
' NB: VarPtr(v) points to the temp VT_BYREF-Variant made by the caller - so we extract the pointer
'
' Prints with address and type:
' 000001D7D0550B70: 000001D7D0550BB0 Long()

Public Function vardump(v, Optional AsBYREF As Boolean) As String
    Dim addr As LongPtr
    Dim typeString As String
    If IsBYREF(v) And Not AsBYREF Then  ' VT_BYREF-Variant --> NonVariant
        addr = read_ptr(VarPtr(v) + 8)
        vardump = hexdumpNonVariant(addr, v)
        typeString = TypeName(v)
    Else ' Variant dump
        addr = VarPtr(v)
        vardump = hexdumpVariant(addr, VarType(v))
        typeString = VarEnumName(v)
    End If
    vardump = hexi_ptr(addr) & ": " & vardump & " " & typeString
End Function

Private Function hexdumpNonVariant(addr As LongPtr, v) As String ' VT_BYREF
    Select Case VarType(v)
        Case vbUserDefinedType  ' mscorlib.Decimal (external Type only in Variant)
            hexdumpNonVariant = dumpaddr(addr, LenB(v)) ' we have no structure info
        Case vbDate, vbCurrency  ' as mem-dump
            hexdumpNonVariant = dumpaddr(addr, 8)
        Case vbDouble ' for VBA32 'vbDouble, vbSingle we print these in FPU style little-endian
            Dim x(1) As Long
            memcpy VarPtr(x(0)), addr, LenB(x(0)) * 2
            hexdumpNonVariant = hexi(x(1), LenB(x(0)) * 2) & hexi(x(0), LenB(x(0)) * 2) ' HI DWORD & LO DWORD
        Case Else      ' pointers, integers: prints little-endian
            Dim d As LongPtr, n_bytes As Long
            n_bytes = sizeof(VarType(v))
            memcpy VarPtr(d), addr, n_bytes ' cast
            hexdumpNonVariant = hexi(d, n_bytes * 2)
    End Select
End Function

Private Function hexdumpVariant(addr As LongPtr, vt As VbVarType) As String
    Select Case vt
        'Case vbUserDefinedType  ' mscorlib.Decimal (external Type only in Variant) print the Variant
        Case vbDecimal   ' special..
            hexdumpVariant = dumpaddr(addr, 16, 2, 1, 1, 4, 4)
        Case vbError
            Dim hresult As Long, fs As String
            memcpy VarPtr(hresult), addr + 8, 4
            If hresult < 0 Then
                fs = "S=1 failure"
            Else
                fs = "S=0 success"
            End If
            Debug.Print fs, "Facility=" & (hresult And FACILITY_MASK) \ &H10000, "Error code=" & (hresult And &HFFFF&)
            hexdumpVariant = dumpaddr(addr, 8 + 2 * LenB(addr), 2, 6, LenB(addr)) ' the Variant
        Case Else
            hexdumpVariant = dumpaddr(addr, 8 + 2 * LenB(addr), 2, 6, LenB(addr)) ' the Variant
    End Select
End Function

Private Function VarEnumName(v) As String
    Dim VTYPE As Integer
    memcpy VarPtr(VTYPE), VarPtr(v), LenB(VTYPE)  ' peek VarType
    Dim fs As String
    fs = flags_string(CLng(VTYPE), strVARENUMFLAGS, 12, 15)
    If vbNullString <> fs Then fs = fs & "|"
    VarEnumName = fs & VarTypeEnumName(VarType(v))   ' TypeName(v) '
End Function

Private Function VarTypeEnumName(vt As Integer) As String ' until 36 vbUserDefinedType
    Const VarEnumNames = "vbEmpty,vbNull,vbInteger,vbLong,vbSingle,vbDuble,vbCurrency,vbDate,vbString,vbObject,vbError,vbBoolean,vbVariant,vbDataObject,vbDecimal,,,vbByte,,,vbLongLong,,,,,,,,,,,,,,,,vbUserDefinedType"
    Dim ty As Long
    ty = vt And VT_TYPEMASK
    If ty > 36 Then
        VarTypeEnumName = "?"
    Else
        VarTypeEnumName = Split(VarEnumNames, ",")(ty)
    End If
End Function


' One line, 16 bytes of mem dump. Little Mid$() opt to avoid too many concat and new BSTR-s

Private Function memdump16(bbuff() As Byte, ByVal firstIdx As Long) As String
    Dim pos As Long, i As Long, b As Byte
    memdump16 = "00 00 00 00 00 00 00 00 | 00 00 00 00 00 00 00 00 | ................"
    pos = 1
    For i = 0 To 15
        b = bbuff(i + firstIdx)
        If b > 0 Then
            If b < 16 Then
                Mid$(memdump16, pos + i * 3 + 1) = Hex$(b)
            Else
                Mid$(memdump16, pos + i * 3) = Hex$(b)
                If b >= 32 Then Mid$(memdump16, i + 53, 1) = Chr$(b)
            End If
        End If
        If i = 7 Then pos = 3
    Next i
End Function

Public Sub memdump(ByVal addr As LongPtr, ByVal rows As Long)
    If rows < 1 Or rows > 128 Then Exit Sub  ' sanity check
    Dim bbuff() As Byte
    read_mem addr, 16 * rows, bbuff ' copy 16 x rows bytes
    Dim idx As Long
    Do
        Debug.Print hexi_ptr(addr + idx) & ": " & memdump16(bbuff, idx)
        idx = idx + 16
        rows = rows - 1
    Loop While rows > 0
    'Debug.Print
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  String

Function print_BSTR(s As String) As String
    If StrPtr(s) Then
        Dim BSTR_Length As Long
        memcpy VarPtr(BSTR_Length), StrPtr(s) - 4, 4
        If LenB(s) Then
            print_BSTR = dumpaddr(StrPtr(s), LenB(s) + 2, LenB(s))
            Dim bbuff() As Byte
            Dim ch As String, i As Long, b As Byte
            ch = String$(LenB(s), ".")
            bbuff = s
            For i = 0 To UBound(bbuff)
                b = bbuff(i)
                If b >= 32 Then Mid$(ch, i + 1, 1) = Chr$(b)
            Next i
            print_BSTR = print_BSTR & " | """ & ch & """"
        Else
            print_BSTR = dumpaddr(StrPtr(s), 2)
        End If
        print_BSTR = hexi_ptr(StrPtr(s) - 4) & ": " & hexi(BSTR_Length, 8) & " | Length" & vbLf & hexi_ptr(StrPtr(s)) & ": " & print_BSTR
    Else
        print_BSTR = "NULL BSTR POINTER"
    End If
End Function

'    Sub test_bstr()
'        Dim s As String
'        Debug.Print print_BSTR("12345")
'        Debug.Print print_BSTR("")
'        Debug.Print print_BSTR(s)
'        s = VBA.vbNullChar
'        Debug.Print print_BSTR(s)
'    End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' VarPtr(arr) is also Compile Error
'
' The only way to get the pointer to SAFEARRAY in VBA is to pass arrays As ByRef Variant
' Will be VT_ARRAY|VT_BYREF type Variant for variables
' Works for both Fix- and Dynamic Arrays
' or.. Variant holding array: VT_ARRAY

' We implement "ArrPtr()" that returns a pointer to SAFEARRAY struct: PSA
' can be null-ptr (before ReDim or after Erase)
' We can pass all types of Arrays as Variant EXCEPT Module UDT Arrays :(

Public Function ArrPtr(a) As LongPtr  ' returns PSA
    If Not IsArray(a) Then ArrPtr = -1: Exit Function
    ArrPtr = read_ptr(VarPtr(a) + 8)
    If IsBYREF(a) Then ArrPtr = read_ptr(ArrPtr) ' ptr-to-SA, can be null
End Function

' shallow-copy of SA with bounds --> into SA and ab()
Private Sub copy_safearray(ByVal psa As LongPtr, ByRef sa As SAFEARRAY, ByRef ab() As SAFEARRAYBOUND)
    memcpy VarPtr(sa), psa, LenB(sa)
    ReDim ab(sa.cDims - 1) ' TODO sanity check?
    memcpy VarPtr(ab(0)), psa + LenB(sa), sa.cDims * LenB(ab(0))
End Sub

Private Sub print_sa(sa As SAFEARRAY, ab() As SAFEARRAYBOUND, have() As LongPtr)
    Debug.Print "SAFEARRAY STRUCTURE:"
    Debug.Print "Offs", "Name", "Value Hex"
    
    If (sa.fFeatures And FADF_HAVEVARTYPE) Then
        Debug.Print -4, "VbVarType", hexi(have(0), 8)  'LenB(ty) * 2)
    ElseIf (sa.fFeatures And FADF_HAVEIID) Then
        Debug.Print -16, "GUID", dumpaddr(VarPtr(have(0)), 16, 4, 2, 2, 2)
    ElseIf (sa.fFeatures And FADF_RECORD) Then ' -4 pointer to the IRecordInfo interface. IRecordInfo interface inherits from the IUnknown interface.
        Debug.Print -LenB(have(0)), "IRecordInfo", hexi(have(0), LenB(have(0)) * 2)
    End If
    
    Debug.Print 0, "cDims", hexi(sa.cDims, 4)
    Debug.Print 2, "fFeatures", hexi(sa.fFeatures, 4) '
    Debug.Print 4, "cbElements", hexi(sa.cbElements, 8)
    Debug.Print 8, "cLocks", hexi(sa.cLocks, 8)
    Debug.Print VarPtr(sa.pvData) - VarPtr(sa), "pvData", hexi_ptr(sa.pvData)
    Dim i As Long, offs As Long
    offs = LenB(sa)
    For i = 0 To sa.cDims - 1
        Debug.Print offs, "cElements", hexi(ab(i).cElements, 8)
        Debug.Print offs + 4, "lLbound", hexi(ab(i).lLbound, 8)
        offs = offs + 8
    Next i
    Dim fs As String
    fs = flags_string(CLng(sa.fFeatures), strADVFEATUREFLAGS, 0, 15)
    If vbNullString = fs Then fs = "NONE"
    Debug.Print "Flags: " & fs
End Sub

Private Sub dump_elements(varr) ' as VT_BYREF Variant: for all array types
    Dim i As Long
    For i = LBound(varr) To UBound(varr)
        Debug.Print vardump(varr(i))
    Next i
    Debug.Print
End Sub

Public Sub dump_safearray(a, Optional withElements As Long)
    Dim psa As LongPtr
    psa = ArrPtr(a)
    If -1 = psa Then
        Debug.Print "Not IsArray"
    ElseIf 0 = psa Then
        Debug.Print "UNallocated SAFEARRAY PSA=0"
    Else
        dump_safearray_psa psa
        If withElements Then dump_elements a
    End If
End Sub

' separate for arrays we cannot pass as Variant: ParamArray, UDT arrays (from memdump)
Public Sub dump_safearray_psa(ByVal psa As LongPtr)
        Dim sa As SAFEARRAY, ab() As SAFEARRAYBOUND ' last dim first
        Dim have(0 To 3) As LongPtr ' Data at -offset. Fits all. works 32/64-bit
        ' shallow-copy of SA to debug
        copy_safearray psa, sa, ab
        
        If (sa.fFeatures And FADF_HAVEVARTYPE) Then
            memcpy VarPtr(have(0)), psa - 4, 4
        ElseIf (sa.fFeatures And FADF_HAVEIID) Then
            memcpy VarPtr(have(0)), psa - 16, 16
        ElseIf (sa.fFeatures And FADF_RECORD) Then ' -4 pointer to the IRecordInfo interface
            memcpy VarPtr(have(0)), psa - LenB(psa), LenB(psa)
        End If
        
        Debug.Print "Addr of SAFEARRAY = " & hexi_ptr(psa)
        print_sa sa, ab, have
        Debug.Print
End Sub

'    Sub test_arr()
'        Dim varr()
'        dump_safearray varr
'        ReDim varr(3)
'        varr(1) = "hi"
'        varr(2) = Array(1, 2, 3)
'        dump_safearray varr, 1
'        ReDim varr(3, 10, 5)
'        dump_safearray varr
'        Debug.Print vardump(varr(0, 0, 0))
'        Debug.Print vardump(varr(3, 10, 5))
'    End Sub


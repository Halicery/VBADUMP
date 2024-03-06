# VBADUMP

Memory dump and layout routines for VBA variables and structures, written in VBA in a Standard Module: **vbadump.bas**. Uses kernel32 RtlMoveMemory to peek. It was written to explore some internal behaviour of VBA, mainly to understand variable storage, parameter passing and especially to explore how Arrays are laid out. This code was used - of course after clearing it up and stripped a little - to write the blog __VBA internal investigations__. 

It is supposed to work for both 32- and 64-bit VBA and it was tested and developed on two versions of Office in Excel: 

- 64-bit Office 365 
- 32-bit Office 2016

## Some main routines: 

### Public Sub memdump(ByVal addr As LongPtr, ByVal rows As Long)

Displays memory content: 

    Dim v
    Dim p As LongPtr
    p = -1
    v = -1
    memdump VarPtr(p), 5
	
	000001D7D0550C28: FF FF FF FF FF FF FF FF | 02 00 00 00 00 00 00 00 | ÿÿÿÿÿÿÿÿ........
	000001D7D0550C38: FF FF 00 00 00 00 00 00 | 00 00 00 00 00 00 00 00 | ÿÿ..............
	000001D7D0550C48: F0 59 7F 56 D8 01 00 00 | F0 59 7F 56 D8 01 00 00 | ðYVØ...ðYVØ...
	000001D7D0550C58: 10 4C F6 CF D7 01 00 00 | 00 00 00 00 00 00 00 00 | .LöÏ×...........
	000001D7D0550C68: 00 00 00 00 00 00 00 00 | 00 00 00 00 00 00 00 00 | ................



### Public Function vardump(v, Optional AsBYREF As Boolean) As String

Dumps memory content of variables:

    Dim v
    ReDim v(3)
    Debug.Print vardump(v)
    
    000001D7D0550C30: 0C20-000000000000-1030A8C9D7010000-0000000000000000 VT_ARRAY|vbVariant

### Public Sub dump_safearray(a, Optional withElements As Long)

Dumps SAFEARRAY structures of different arrays: 

    ReDim sarr(3) As String
    sarr(0) = "hi"
    sarr(1) = ""
    dump_safearray sarr, 1

    Addr of SAFEARRAY = 000001D7CC2B25B0
    SAFEARRAY STRUCTURE:
    Offs          Name          Value Hex
    -4            VbVarType     00000008
     0            cDims         0001
     2            fFeatures     0180
     4            cbElements    00000008
     8            cLocks        00000000
     16           pvData        000001D7D81FAC80
     24           cElements     00000004
     28           lLbound       00000000
    Flags: FADF_BSTR|FADF_HAVEVARTYPE
    
    000001D7D81FAC80: 000001D7D81FABC8 String
    000001D7D81FAC88: 000001D7D81FAB98 String
    000001D7D81FAC90: 0000000000000000 String
    000001D7D81FAC98: 0000000000000000 String








Attribute VB_Name = "mStrings"
Option Explicit

Private Declare Function PutMem4 Lib "msvbvm60.dll" (ByVal Addr As Long, ByVal NewVal As Long) As Long
Private Declare Function SysAllocString Lib "oleaut32.dll" (Optional ByVal pszStrPtr As Long) As Long

Public Function stringGetFromPointerW(ByVal p_StringPointer As Long) As String
   ' This code courtesy of Bonnie West @
   ' http://www.vbforums.com/showthread.php?707879-VB6-Dereferencing-Pointers-sans-CopyMemory&p=4330835&viewfull=1#post4330835
   
   PutMem4 VarPtr(stringGetFromPointerW), SysAllocString(p_StringPointer)
End Function


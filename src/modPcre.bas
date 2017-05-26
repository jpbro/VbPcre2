Attribute VB_Name = "modPcre"
Option Explicit

' MIT License
'
' Copyright (c) 2017 Jason Peter Brown
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

Private Type pcreCalloutBlock
   Version As Long
   CalloutNumber As Long
   CaptureTop As Long
   CaptureLast As Long
   OffsetVectorPointer As Long
   MarkPointer As Long
   SubjectPointer As Long
   SubjectLength As Long
   StartMatch As Long
   CurrentPosition As Long
   PatternPosition As Long
   NextItemLength As Long
   CalloutStringOffset As Long
   CalloutStringLength As Long
   CalloutStringPointer As Long
End Type

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Public Function pcreCalloutProc(ByVal p_CalloutBlockPointer As Long, ByRef p_UserData As Long) As Long
   ' RETURN VALUES FROM CALLOUTS
   ' The external callout function returns an integer to PCRE2.
   ' If the value is zero, matching proceeds as normal.
   ' If the value is greater than zero, matching fails at the current point, but the testing of other matching possibilities goes ahead, just as if a lookahead assertion had failed.
   ' If the value is less than zero, the match is abandoned, and the matching function returns the negative value.
   ' Negative values should normally be chosen from the set of PCRE2_ERROR_xxx values.
   ' In particular, PCRE2_ERROR_NOMATCH forces a standard "no match" failure.
   ' The error number PCRE2_ERROR_CALLOUT is reserved for use by callout functions; it will never be used by PCRE2 itself.
   
   Dim lt_CalloutBlock As modPcre.pcreCalloutBlock
   
   MsgBox "INCALLOUT"
   Debug.Assert False
   Debug.Print "IN CALLOUT"
   
   CopyMemory ByVal VarPtr(lt_CalloutBlock), p_CalloutBlockPointer, LenB(lt_CalloutBlock)

   Debug.Print lt_CalloutBlock.Version

   Debug.Print "OUT CALLOUT"
End Function


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

Public Type pcreCalloutBlock
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

Public Type pcreCalloutEnumerateBlock
   Version As Long
   PatternPosition As Long
   NextItemLength As Long
   CalloutNumber As Long
   CalloutStringOffset As Long
   CalloutStringLength As Long
   CalloutStringPointer As Long
End Type

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Public Function pcreCalloutEnumerateProc(ByVal p_CalloutEnumerateBlockPointer As Long, ByRef p_UserData As Long) As Long
   ' RETURN VALUES FROM CALLOUTS
   ' The external callout function returns an integer to PCRE2.
   ' If the value is zero, matching proceeds as normal.
   ' If the value is greater than zero, matching fails at the current point, but the testing of other matching possibilities goes ahead, just as if a lookahead assertion had failed.
   ' If the value is less than zero, the match is abandoned, and the matching function returns the negative value.
   ' Negative values should normally be chosen from the set of PCRE2_ERROR_xxx values.
   ' In particular, PCRE2_ERROR_NOMATCH forces a standard "no match" failure.
   ' The error number PCRE2_ERROR_CALLOUT is reserved for use by callout functions; it will never be used by PCRE2 itself.
   
   Dim lt_CalloutEnumerateBlock As modPcre.pcreCalloutEnumerateBlock
   Dim lo_Pcre As CPcre
   
   Debug.Print "In pcreCalloutEnumerateProc"
   
   ' Get a weak reference to the appropriate PCRE object
   Set lo_Pcre = GetWeakReference(p_UserData)
   
   CopyMemory ByVal VarPtr(lt_CalloutEnumerateBlock), ByVal p_CalloutEnumerateBlockPointer, LenB(lt_CalloutEnumerateBlock)

   ' Ask the PCRE object to raise an event so the hosting code can respond to the callout
   pcreCalloutEnumerateProc = lo_Pcre.RaiseCalloutEnumeratedEvent(lt_CalloutEnumerateBlock)

   Debug.Print "Out pcreCalloutEnumerateProc. Result: " & pcreCalloutEnumerateProc
End Function

Private Function GetWeakReference(ByVal p_Pointer As Long) As CPcre
   ' Can't remember where I found this code a long time ago -
   ' would be very happy to credit the originator if anyone knows who it is?
   
   Dim lo_Object As CPcre
   
   CopyMemory lo_Object, p_Pointer, 4&

   Set GetWeakReference = lo_Object

   CopyMemory lo_Object, 0&, 4&
End Function


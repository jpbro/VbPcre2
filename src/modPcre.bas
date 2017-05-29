Attribute VB_Name = "modPcre"
Option Explicit

' Copyright (c) 2017 Jason Peter Brown
'
' MIT License
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

Public Enum PCRE_ReturnCode
    PCRE_RC_OK = 0
    
    'Error codes: no match and partial match are "expected" errors.
    PCRE_RC_ERROR_NOMATCH = -1
    PCRE_RC_ERROR_PARTIAL = -2
    
    'Error codes for UTF-8 validity checks
    PCRE_RC_ERROR_UTF8_ERR1 = -3
    PCRE_RC_ERROR_UTF8_ERR2 = -4
    PCRE_RC_ERROR_UTF8_ERR3 = -5
    PCRE_RC_ERROR_UTF8_ERR4 = -6
    PCRE_RC_ERROR_UTF8_ERR5 = -7
    PCRE_RC_ERROR_UTF8_ERR6 = -8
    PCRE_RC_ERROR_UTF8_ERR7 = -9
    PCRE_RC_ERROR_UTF8_ERR8 = -10
    PCRE_RC_ERROR_UTF8_ERR9 = -11
    PCRE_RC_ERROR_UTF8_ERR10 = -12
    PCRE_RC_ERROR_UTF8_ERR11 = -13
    PCRE_RC_ERROR_UTF8_ERR12 = -14
    PCRE_RC_ERROR_UTF8_ERR13 = -15
    PCRE_RC_ERROR_UTF8_ERR14 = -16
    PCRE_RC_ERROR_UTF8_ERR15 = -17
    PCRE_RC_ERROR_UTF8_ERR16 = -18
    PCRE_RC_ERROR_UTF8_ERR17 = -19
    PCRE_RC_ERROR_UTF8_ERR18 = -20
    PCRE_RC_ERROR_UTF8_ERR19 = -21
    PCRE_RC_ERROR_UTF8_ERR20 = -22
    PCRE_RC_ERROR_UTF8_ERR21 = -23

    'Error codes for UTF-16 validity checks
    PCRE_RC_ERROR_UTF16_ERR1 = -24
    PCRE_RC_ERROR_UTF16_ERR2 = -25
    PCRE_RC_ERROR_UTF16_ERR3 = -26

    'Error codes for UTF-32 validity checks
    PCRE_RC_ERROR_UTF32_ERR1 = -27
    PCRE_RC_ERROR_UTF32_ERR2 = -28

    'Error codes for pcre2[_dfa]_match= , substring extraction functions, context
    ' functions, and serializing functions. They are in numerical order. Originally
    ' they were in alphabetical order too, but now that PCRE2 is released, the
    ' numbers must not be changed.
    PCRE_RC_ERROR_BADDATA = -29
    PCRE_RC_ERROR_MIXEDTABLES = -30        ' Name was changed
    PCRE_RC_ERROR_BADMAGIC = -31
    PCRE_RC_ERROR_BADMODE = -32
    PCRE_RC_ERROR_BADOFFSET = -33
    PCRE_RC_ERROR_BADOPTION = -34
    PCRE_RC_ERROR_BADREPLACEMENT = -35
    PCRE_RC_ERROR_BADUTFOFFSET = -36
    PCRE_RC_ERROR_CALLOUT = -37            ' Never used by PCRE2 itself
    PCRE_RC_ERROR_DFA_BADRESTART = -38
    PCRE_RC_ERROR_DFA_RECURSE = -39
    PCRE_RC_ERROR_DFA_UCOND = -40
    PCRE_RC_ERROR_DFA_UFUNC = -41
    PCRE_RC_ERROR_DFA_UITEM = -42
    PCRE_RC_ERROR_DFA_WSSIZE = -43
    PCRE_RC_ERROR_INTERNAL = -44
    PCRE_RC_ERROR_JIT_BADOPTION = -45
    PCRE_RC_ERROR_JIT_STACKLIMIT = -46
    PCRE_RC_ERROR_MATCHLIMIT = -47
    PCRE_RC_ERROR_NOMEMORY = -48
    PCRE_RC_ERROR_NOSUBSTRING = -49
    PCRE_RC_ERROR_NOUNIQUESUBSTRING = -50
    PCRE_RC_ERROR_NULL = -51
    PCRE_RC_ERROR_RECURSELOOP = -52
    PCRE_RC_ERROR_RECURSIONLIMIT = -53
    PCRE_RC_ERROR_UNAVAILABLE = -54
    PCRE_RC_ERROR_UNSET = -55
    PCRE_RC_ERROR_BADOFFSETLIMIT = -56
    PCRE_RC_ERROR_BADREPESCAPE = -57
    PCRE_RC_ERROR_REPMISSINGBRACE = -58
    PCRE_RC_ERROR_BADSUBSTITUTION = -59
    PCRE_RC_ERROR_BADSUBSPATTERN = -60
    PCRE_RC_ERROR_TOOMANYREPLACE = -61
    PCRE_RC_ERROR_BADSERIALIZEDDATA = -62
End Enum

'The following option bits can be passed only to pcre2_compile(). However,
' they may affect compilation, JIT compilation, and/or interpretive execution.
' The following tags indicate which:
'
' C   alters what is compiled by pcre2_compile()
' J   alters what is compiled by pcre2_jit_compile()
' M   is inspected during pcre2_match() execution
' D   is inspected during pcre2_dfa_match() execution
Public Enum PCRE_CompileOptions
    PCRE_CO_ALLOW_EMPTY_CLASS = &H1&           ' C
    PCRE_CO_ALT_BSUX = &H2&                    ' C
    PCRE_CO_AUTO_CALLOUT = &H4&                ' C
    PCRE_CO_CASELESS = &H8&                    ' C
    PCRE_CO_DOLLAR_ENDONLY = &H10&             '   J M D
    PCRE_CO_DOTALL = &H20&                     ' C
    PCRE_CO_DUPNAMES = &H40&                   ' C
    PCRE_CO_EXTENDED = &H80&                   ' C
    PCRE_CO_FIRSTLINE = &H100&                 '   J M D
    PCRE_CO_MATCH_UNSET_BACKREF = &H200&       ' C J M
    PCRE_CO_MULTILINE = &H400&                 ' C
    PCRE_CO_NEVER_UCP = &H800&                 ' C
    PCRE_CO_NEVER_UTF = &H1000&                ' C
    PCRE_CO_NO_AUTO_CAPTURE = &H2000&          ' C
    PCRE_CO_NO_AUTO_POSSESS = &H4000&          ' C
    PCRE_CO_NO_DOTSTAR_ANCHOR = &H8000&        ' C
    PCRE_CO_NO_START_OPTIMIZE = &H10000        '   J M D
    PCRE_CO_UCP = &H20000                      ' C J M D
    PCRE_CO_UNGREEDY = &H40000                 ' C
    PCRE_CO_UTF = &H80000                      ' C J M D
    PCRE_CO_NEVER_BACKSLASH_C = &H100000       ' C
    PCRE_CO_ALT_CIRCUMFLEX = &H200000          '   J M D
    PCRE_CO_ALT_VERBNAMES = &H400000           ' C
    PCRE_CO_USE_OFFSET_LIMIT = &H800000        '   J M D
End Enum

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


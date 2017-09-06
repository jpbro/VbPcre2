Attribute VB_Name = "modPcre2"
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
   PCRE_RC_ERROR_MIXEDTABLES = -30         ' Name was changed
   PCRE_RC_ERROR_BADMAGIC = -31
   PCRE_RC_ERROR_BADMODE = -32
   PCRE_RC_ERROR_BADOFFSET = -33
   PCRE_RC_ERROR_BADOPTION = -34
   PCRE_RC_ERROR_BADREPLACEMENT = -35
   PCRE_RC_ERROR_BADUTFOFFSET = -36
   PCRE_RC_ERROR_CALLOUT = -37             ' Never used by PCRE2 itself
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

   [_PCRE_RC_ERROR_FIRST] = -1
   [_PCRE_RC_ERROR_LAST] = -62   ' If you add more PCRE2 error codes, make sure to update this value!
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
   PCRE_CO_ALLOW_EMPTY_CLASS = &H1&            ' C
   PCRE_CO_ALT_BSUX = &H2&                     ' C
   PCRE_CO_AUTO_CALLOUT = &H4&                 ' C
   PCRE_CO_CASELESS = &H8&                     ' C
   PCRE_CO_DOLLAR_ENDONLY = &H10&              '   J M D
   PCRE_CO_DOTALL = &H20&                      ' C
   PCRE_CO_DUPNAMES = &H40&                    ' C
   PCRE_CO_EXTENDED = &H80&                    ' C
   PCRE_CO_FIRSTLINE = &H100&                  '   J M D
   PCRE_CO_MATCH_UNSET_BACKREF = &H200&        ' C J M
   PCRE_CO_MULTILINE = &H400&                  ' C
   PCRE_CO_NEVER_UCP = &H800&                  ' C
   PCRE_CO_NEVER_UTF = &H1000&                 ' C
   PCRE_CO_NO_AUTO_CAPTURE = &H2000&           ' C
   PCRE_CO_NO_AUTO_POSSESS = &H4000&           ' C
   PCRE_CO_NO_DOTSTAR_ANCHOR = &H8000&         ' C
   PCRE_CO_NO_START_OPTIMIZE = &H10000         '   J M D
   PCRE_CO_UCP = &H20000                       ' C J M D
   PCRE_CO_UNGREEDY = &H40000                  ' C
   PCRE_CO_UTF = &H80000                       ' C J M D
   PCRE_CO_NEVER_BACKSLASH_C = &H100000        ' C
   PCRE_CO_ALT_CIRCUMFLEX = &H200000           '   J M D
   PCRE_CO_ALT_VERBNAMES = &H400000            ' C
   PCRE_CO_USE_OFFSET_LIMIT = &H800000         '   J M D
End Enum

Public Const PCRE2_ANCHORED             As Long = &H80000000
Public Const PCRE2_NO_UTF_CHECK        As Long = &H40000000
Public Const PCRE2_NOTBOL As Long = &H1
Public Const PCRE2_NOTEOL As Long = &H2
Public Const PCRE2_NOTEMPTY  As Long = &H4
Public Const PCRE2_NOTEMPTY_ATSTART As Long = &H8
Public Const PCRE2_PARTIAL_SOFT As Long = &H10
Public Const PCRE2_PARTIAL_HARD As Long = &H20
Public Const PCRE2_ERROR_NOMATCH As Long = -1
Public Const PCRE2_SUBSTITUTE_GLOBAL As Long = &H100

Public Declare Function pcre2_compile_context_create Lib "pcre2-16.dll" Alias "_pcre2_compile_context_create_16@4" (Optional ByVal p_MallocFunc As Long = 0&) As Long
Public Declare Sub pcre2_compile_context_free Lib "pcre2-16.dll" Alias "_pcre2_compile_context_free_16@4" (ByVal p_ContextHandle As Long)
Public Declare Function pcre2_compile Lib "pcre2-16.dll" Alias "_pcre2_compile_16@24" (ByVal p_RegexStringPointer As Long, ByVal p_RegexStringLength As Long, ByVal p_CompileOptions As PCRE_CompileOptions, ByRef p_ErrorCode As PCRE_ReturnCode, ByRef p_CharWhereErrorOccured As Long, Optional ByVal p_CompileContextHandle As Long = &H0) As Long
Public Declare Sub pcre2_code_free Lib "pcre2-16.dll" Alias "_pcre2_code_free_16@4" (ByVal p_CompiledRegecHandle As Long)
Public Declare Function pcre2_match_data_create_from_pattern Lib "pcre2-16.dll" Alias "_pcre2_match_data_create_from_pattern_16@8" (ByVal p_CompiledRegexHandle As Long, ByVal p_Options As Long) As Long
Public Declare Function pcre2_match Lib "pcre2-16.dll" Alias "_pcre2_match_16@28" (ByVal p_CompiledRegexHandle As Long, ByVal p_StringToSearchPointer As Long, ByVal p_StringToSearchLength As Long, ByVal p_StartSearchOffset As Long, ByVal p_MatchOptions As Long, ByVal p_MatchDataHandle As Long, ByVal p_MatchContextHandle As Long) As Long
Public Declare Function pcre2_get_ovector_pointer Lib "pcre2-16.dll" Alias "_pcre2_get_ovector_pointer_16@4" (ByVal p_MatchDataHandle As Long) As Long
Public Declare Sub pcre2_match_data_free Lib "pcre2-16.dll" Alias "_pcre2_match_data_free_16@4" (ByVal p_MatchDataHandle As Long)
Public Declare Function pcre2_callout_enumerate Lib "pcre2-16.dll" Alias "_pcre2_callout_enumerate_16@12" (ByVal p_CompiledRegexHandle As Long, ByVal p_CalloutAddress As Long, ByVal p_CalloutDataPointer As Long) As Long
Public Declare Function pcre2_set_callout Lib "pcre2-16.dll" Alias "_pcre2_set_callout_16@12" (ByVal p_MatchContextHandle As Long, ByVal p_CalloutAddress As Long, ByVal p_CalloutDataPointer As Long) As Long
Public Declare Function pcre2_substitute Lib "pcre2-16.dll" Alias "_pcre2_substitute_16@44" (ByVal p_CompiledRegexHandle As Long, ByVal p_StringToSearchPointer As Long, ByVal p_StringToSearchLength As Long, ByVal p_StartSearchOffset As Long, ByVal p_MatchOptions As Long, ByVal p_MatchDataHandle As Long, ByVal p_MatchContextHandle As Long, ByVal p_ReplacementTextPointer As Long, ByVal p_ReplacementTextLength As Long, ByVal p_OutputBufferPointer As Long, ByRef p_OutputBufferLength As Long) As Long
Public Declare Function pcre2_match_context_create Lib "pcre2-16.dll" Alias "_pcre2_match_context_create_16@4" (ByVal p_GeneralContext As Long) As Long
Public Declare Function pcre2_match_context_free Lib "pcre2-16.dll" Alias "_pcre2_match_context_free_16@4" (ByVal p_MatchContextHandle As Long) As Long
Public Declare Function pcre2_get_ovector_count Lib "pcre2-16.dll" Alias "_pcre2_get_ovector_count_16@4" (ByVal p_MatchDataHandle As Long) As Long
Public Declare Function pcre2_get_error_message Lib "pcre2-16.dll" Alias "_pcre2_get_error_message_16@12" (ByVal p_ErrorCode As Long, ByVal p_ErrorMessageBufferPointer As Long, ByVal p_ErrorMessageBufferLength As Long) As Long

Public Function pcreCalloutProc(ByVal p_CalloutBlockPointer As Long, ByVal p_UserData As Long) As Long
   Dim lt_CalloutBlock As modPcre2.pcreCalloutBlock
   Dim lo_Pcre As cPcre2
   
   Debug.Print "In pcreCalloutProc"
   Debug.Print "Recevied callout from ObjPtr: " & p_UserData

   ' Get a weak reference to the appropriate PCRE object
   If p_UserData = 0 Then
      ' Should be ObjPtr of your CPcre2 object!
      Debug.Assert False
      
   Else
      Set lo_Pcre = GetWeakReference(p_UserData)
      
      win32_CopyMemory ByVal VarPtr(lt_CalloutBlock), ByVal p_CalloutBlockPointer, LenB(lt_CalloutBlock)
   
      ' Ask the PCRE object to raise an event so the hosting code can respond to the callout
      pcreCalloutProc = lo_Pcre.RaiseCalloutReceivedEvent(lt_CalloutBlock)
   End If
   
   Debug.Print "Out pcreCalloutProc"
End Function

Public Function pcreCalloutEnumerateProc(ByVal p_CalloutEnumerateBlockPointer As Long, ByVal p_UserData As Long) As Long
   ' RETURN VALUES FROM CALLOUTS
   ' The external callout function returns an integer to PCRE2.
   ' If the value is zero, matching proceeds as normal.
   ' If the value is greater than zero, matching fails at the current point, but the testing of other matching possibilities goes ahead, just as if a lookahead assertion had failed.
   ' If the value is less than zero, the match is abandoned, and the matching function returns the negative value.
   ' Negative values should normally be chosen from the set of PCRE2_ERROR_xxx values.
   ' In particular, PCRE2_ERROR_NOMATCH forces a standard "no match" failure.
   ' The error number PCRE2_ERROR_CALLOUT is reserved for use by callout functions; it will never be used by PCRE2 itself.
   
   Dim lt_CalloutEnumerateBlock As modPcre2.pcreCalloutEnumerateBlock
   Dim lo_Pcre As cPcre2
   
   Debug.Print "In pcreCalloutEnumerateProc"
   
   If p_UserData = 0 Then
      ' Should be ObjPtr of your CPcre2 object!
      Debug.Assert False
      
   Else
      ' Get a weak reference to the appropriate PCRE object
      Set lo_Pcre = GetWeakReference(p_UserData)
      
      win32_CopyMemory ByVal VarPtr(lt_CalloutEnumerateBlock), ByVal p_CalloutEnumerateBlockPointer, LenB(lt_CalloutEnumerateBlock)
   
      ' Ask the PCRE object to raise an event so the hosting code can respond to the callout
      pcreCalloutEnumerateProc = lo_Pcre.RaiseCalloutEnumeratedEvent(lt_CalloutEnumerateBlock)
   End If
   
   Debug.Print "Out pcreCalloutEnumerateProc. Result: " & pcreCalloutEnumerateProc
End Function

Private Function GetWeakReference(ByVal p_Pointer As Long) As cPcre2
   ' Can't remember where I found this code a long time ago -
   ' would be very happy to credit the originator if anyone knows who it is?
   
   Dim lo_Object As cPcre2
   
   win32_CopyMemory lo_Object, p_Pointer, 4&

   Set GetWeakReference = lo_Object

   win32_CopyMemory lo_Object, 0&, 4&
End Function


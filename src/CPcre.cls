VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPcre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Public Enum e_EnumerateCalloutAction
   enumcalloutaction_Continue = 0
   enumcalloutaction_StopEnumeration = 1
   enumcalloutaction_StopAll = 2

   enumcalloutaction_RaiseError = vbObjectError
End Enum


Public Event CalloutEnumerated(ByVal p_CalloutNumber As Long, _
                               ByVal p_PatternPosition As Long, _
                               ByVal p_NextItemLength As Long, _
                               ByVal p_CalloutOffset As Long, _
                               ByVal p_CalloutLength As Long, _
                               ByVal p_CalloutString As String, _
                               ByRef p_Action As e_EnumerateCalloutAction)

Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal p_LibraryFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal p_Hmodule As Long) As Long

Private Declare Function pcre2_compile_context_create Lib "pcre2-16.dll" Alias "_pcre2_compile_context_create_16@4" (Optional ByVal p_MallocFunc As Long = 0&) As Long
Private Declare Sub pcre2_compile_context_free Lib "pcre2-16.dll" Alias "_pcre2_compile_context_free_16@4" (ByVal p_ContextHandle As Long)
Private Declare Function pcre2_compile Lib "pcre2-16.dll" Alias "_pcre2_compile_16@24" (ByVal p_RegexStringPointer As Long, ByVal p_RegexStringLength As Long, ByVal p_CompileOptions As Long, ByRef p_ErrorCode As Long, ByRef p_CharWhereErrorOccured As Long, Optional ByVal p_ContextHandle As Long = &H0) As Long
Private Declare Sub pcre2_code_free Lib "pcre2-16.dll" Alias "_pcre2_code_free_16@4" (ByVal p_CompiledRegecHandle As Long)
Private Declare Function pcre2_match_data_create_from_pattern Lib "pcre2-16.dll" Alias "_pcre2_match_data_create_from_pattern_16@8" (ByVal p_CompiledRegexHandle As Long, ByVal p_Options As Long) As Long
Private Declare Function pcre2_match Lib "pcre2-16.dll" Alias "_pcre2_match_16@28" (ByVal p_CompiledRegexHandle As Long, ByVal p_StringToSearchPointer As Long, ByVal p_StringToSearchLength As Long, ByVal p_StartSearchOffset As Long, ByVal p_MatchOptions As Long, ByVal p_MatchDataHandle As Long, ByVal p_ContextHandle As Long) As Long
Private Declare Function pcre2_get_ovector_pointer Lib "pcre2-16.dll" Alias "_pcre2_get_ovector_pointer_16@4" (ByVal p_MatchDataHandle As Long) As Long
Private Declare Sub pcre2_match_data_free Lib "pcre2-16.dll" Alias "_pcre2_match_data_free_16@4" (ByVal p_MatchDataHandle As Long)
Private Declare Function pcre2_set_callout Lib "pcre2-16.dll" Alias "_pcre2_set_callout_16@12" (ByVal p_MatchDataHandle As Long, ByVal p_CalloutAddress As Long, ByRef p_CalloutDataPointer As Long) As Long
Private Declare Function pcre2_callout_enumerate Lib "pcre2-16.dll" Alias "_pcre2_callout_enumerate_16@12" (ByVal p_CompiledRegexHandle As Long, ByVal p_CalloutAddress As Long, ByRef p_CalloutDataPointer As Long) As Long

Private Const PCRE2_ERROR_NOMATCH As Long = -1

Private m_Hlib As Long

' Options
Private mo_RegexOptions As CPcreRegexOptions

' Cached Handles
Private m_CompiledRegex As Long
Private m_Context As Long

Private m_Pattern As String   ' Last used regex pattern

Public Function RegexOptions() As CPcreRegexOptions
   If mo_RegexOptions Is Nothing Then Set mo_RegexOptions = New CPcreRegexOptions
   
   Set RegexOptions = mo_RegexOptions
End Function

Private Sub Class_Initialize()
   m_Hlib = LoadLibrary("pcre2-16.dll")
   
   If m_Hlib = 0 Then Err.Raise vbObjectError, , "Could not initialize PCRE2 library! Last DLL Error: " & Err.LastDllError
End Sub

Private Sub Class_Terminate()
   Cleanup
   
   If m_Hlib <> 0 Then
      FreeLibrary m_Hlib
   End If
End Sub

Public Function Execute(ByVal p_TextToSearch As String, Optional ByVal p_RegexToMatch As String) As CPcreMatches
   Static s_LastRegex As String
   Static s_LastRegexOptions As Long
   
   ' Returns CPcreMatches object. If MatchCount = 0 then no matches.

   Dim l_Context As Long
   Dim l_CompiledRegex As Long
   Dim l_ErrorCode As Long
   Dim l_ErrorPosition As Long
   Dim l_ErrorNumber As Long
   Dim l_ErrorDesc As String
   Dim l_Recompile As Boolean
   Dim l_MatchData As Long
   Dim l_MatchCount As Long
   Dim l_OvectorPtr As Long
   Dim la_Ovector() As Long
   Dim l_GlobalSearch As Boolean
   Dim l_EnumerateResult As e_EnumerateCalloutAction
   
   ' Use the value of the Pattern property if option p_RegexToMatch parameter is empty
   ' Otherwise set the value of the Pattern property to the passed p_RegexToMatch value
   If p_RegexToMatch = "" Then
      p_RegexToMatch = Me.Pattern
   Else
      Me.Pattern = p_RegexToMatch
   End If
   
   ' Check if we should recompile regex
   If m_Context = 0 Then
      l_Recompile = True
   Else
      If m_CompiledRegex = 0 Then
         l_Recompile = True
      Else
         If Me.RegexOptions.OptionFlags <> s_LastRegexOptions Then
            l_Recompile = True
         Else
            If Len(s_LastRegex) <> Len(p_RegexToMatch) Then
               l_Recompile = True
            Else
               If s_LastRegex <> p_RegexToMatch Then
                  l_Recompile = True
               End If
            End If
         End If
      End If
   End If
   
   If l_Recompile Then
      'Compile a source string.  Because we are using the 16-bit version, we can just pass StrPtr
      
      Cleanup
      
      l_Context = pcre2_compile_context_create(0)
      If l_Context = 0 Then Err.Raise "Could not compile PCRE context! Last DLL Error: " & Err.LastDllError
      
      m_Context = l_Context
      
      l_CompiledRegex = pcre2_compile(StrPtr(p_RegexToMatch), Len(p_RegexToMatch), Me.RegexOptions.OptionFlags, l_ErrorCode, l_ErrorPosition, l_Context)
      
      If l_CompiledRegex = 0 Then Err.Raise vbObjectError, , "Could not compile regex! Regex: " & p_RegexToMatch & vbNewLine & "Errorcode: " & l_ErrorCode & ", Error Position: " & l_ErrorPosition
      
      s_LastRegexOptions = Me.RegexOptions.OptionFlags
      s_LastRegex = p_RegexToMatch
      m_CompiledRegex = l_CompiledRegex
      
   Else
      ' Use cached regex and context handles
      l_CompiledRegex = m_CompiledRegex
      l_Context = m_Context
      
   End If

   ' Create match data space
   l_MatchData = pcre2_match_data_create_from_pattern(l_CompiledRegex, 0)
   
   If l_MatchData = 0 Then Err.Raise vbObjectError, , "Could not allocate match data! Last DLL Error: " & Err.LastDllError

   If Me.RegexOptions.EnumerateCallouts Then
      ' Enumerate all callouts before matching
      
      l_EnumerateResult = pcre2_callout_enumerate(l_CompiledRegex, AddressOf pcreCalloutEnumerateProc, ObjPtr(Me))
      
      Select Case l_EnumerateResult
      Case enumcalloutaction_StopAll
         ' Cancel matching.
         Debug.Assert False
         GoTo NoErrorCleanup
         
      Case Is < 0
         ' Raise error
         Err.Raise l_EnumerateResult, , "Host reported error."
         
      End Select
   End If
   
   l_GlobalSearch = Me.RegexOptions.GlobalSearch
   
   Set Execute = New CPcreMatches
   
   Do
      l_MatchCount = pcre2_match(l_CompiledRegex, StrPtr(p_TextToSearch), Len(p_TextToSearch), 0, 0, l_MatchData, 0)
   
      Select Case l_MatchCount
      Case PCRE2_ERROR_NOMATCH
         ' No matches, that's normal :)
      
      Case Is > 0
         ' Number of matches, store information about matches
         l_OvectorPtr = pcre2_get_ovector_pointer(l_MatchData)
         
         If l_OvectorPtr = 0 Then
            ' Shouldn't happen!
            Err.Raise vbObjectError, , "Ovector pointer could not be retrieved!"
         End If
            
         Dim l_MatchLen As Long
         
         CopyMemory l_MatchLen, ByVal (l_OvectorPtr + l_MatchCount * 2), 4
         
         Execute.Add Left$(p_TextToSearch, l_MatchLen), l_OvectorPtr, l_MatchCount
      
         If l_GlobalSearch Then
            p_TextToSearch = Mid$(p_TextToSearch, l_MatchLen + 1)
         End If
      
      Case Else
         ' Uhoh! We need to handle these
         Err.Raise vbObjectError - l_MatchCount, , "PCRE Match Error: " & l_MatchCount
      End Select
      
   Loop While l_GlobalSearch And (l_MatchCount > 0)
   
NoErrorCleanup:
   On Error Resume Next
   
   l_ErrorNumber = 0
   l_ErrorDesc = ""
   
Cleanup:
   On Error Resume Next

   ' Free match data if necessary
   If l_MatchData <> 0 Then pcre2_match_data_free l_MatchData: l_MatchData = 0

   If l_ErrorNumber <> 0 Then
      Cleanup
      
      On Error GoTo 0
      Err.Raise l_ErrorNumber, , l_ErrorDesc
   End If

   Exit Function

ErrorHandler:
   Debug.Assert False
   
   l_ErrorNumber = Err.Number
   l_ErrorDesc = Err.Description
      
   Resume Cleanup
End Function

Private Sub Cleanup()
   'Free compiled regex before exiting
   If m_CompiledRegex <> 0 Then pcre2_code_free m_CompiledRegex: m_CompiledRegex = 0

   'Free compile context before exiting
   If m_Context <> 0 Then pcre2_compile_context_free m_Context: m_Context = 0
End Sub

Friend Function RaiseCalloutEnumeratedEvent(pt_CalloutEnumerateBlock As pcreCalloutEnumerateBlock) As Long
   Dim l_CalloutString As String
   Dim l_Action As e_EnumerateCalloutAction
   
   With pt_CalloutEnumerateBlock
      If .CalloutStringPointer <> 0 Then
         l_CalloutString = stringGetFromPointerW(.CalloutStringPointer)
      End If
      
      RaiseEvent CalloutEnumerated(.CalloutNumber, .PatternPosition, .NextItemLength, .CalloutStringOffset, .CalloutStringLength, l_CalloutString, l_Action)
   End With
   
   RaiseCalloutEnumeratedEvent = l_Action
End Function

' GlobalSearch and IgnoreCase properties are aliases for GlobalSearch and CaseSensitive options
' To more closely match VBScript Regex Object Model. Unfortunately we can't have a property call Global
' because it is a reserved keyword, so we can match VBScript Regex OM exactly :(
Public Property Let GlobalSearch(ByVal p_Enable As Boolean)
   Me.RegexOptions.GlobalSearch = p_Enable
End Property

Public Property Get GlobalSearch() As Boolean
   GlobalSearch = Me.RegexOptions.GlobalSearch
End Property

Public Property Let IgnoreCase(ByVal p_Enable As Boolean)
   Me.RegexOptions.CaseSensitive = Not p_Enable
End Property

Public Property Get IgnoreCase() As Boolean
   IgnoreCase = Not Me.RegexOptions.CaseSensitive
End Property

' Pattern property added to more closely mimic VBScript Regex object model
Public Property Get Pattern() As String
   Pattern = m_Pattern
End Property

Public Property Let Pattern(ByVal p_RegexPattern As String)
   m_Pattern = p_RegexPattern
End Property

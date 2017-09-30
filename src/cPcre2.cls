VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPcre2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Copyright (c) 2017 Jason Peter Brown <jason@bitspaces.com>
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

Public Enum e_CalloutEnumeratedAction
   enumcalloutaction_Continue = 0
   enumcalloutaction_StopEnumeration = 1
   enumcalloutaction_StopAll = 2

   enumcalloutaction_RaiseError = vbObjectError
End Enum

Public Enum e_CalloutReceivedAction
   enumcalloutaction_Continue = 0   ' Continue normally
   enumcalloutaction_Skip = 1 ' Skip the current match, but continue matching.
   enumcalloutaction_Cancel = -1 ' Any negative number will cancel matching and return the passed number
End Enum

Public Enum e_SubstitutionAction
   subaction_None ' Perform no substitution
   subaction_Replace ' Replace the match with the supplied text
   subaction_ReplaceAndCache  ' Replace the match with the supplied text. If a subsequent exact match is found, replace with the same cached value and skip the Matched event.
End Enum

Public Event Matched(ByRef p_MatchedText As String, ByRef p_SubstitutionAction As e_SubstitutionAction, ByRef p_Cancel As Boolean)   ' If GlobalSearch = True and EnableMatchedEvent = True then this event will be raised for each match.
                                                                                 ' You can change p_MatchedText and set p_Replace = True to perform a substitution.

Public Event CalloutEnumerated(ByVal p_CalloutNumber As Long, _
                               ByVal p_CalloutLabel As String, _
                               ByVal p_CalloutOffset As Long, _
                               ByVal p_PatternPosition As Long, _
                               ByVal p_NextItemLength As Long, _
                               ByRef p_Action As e_CalloutEnumeratedAction)

Public Event CalloutReceived(ByVal p_CalloutNumber As Long, _
                             ByVal p_CalloutLabel As String, _
                             ByVal p_CalloutOffset As Long, _
                             ByVal p_Subject As String, _
                             ByVal p_Mark As String, _
                             ByVal p_CaptureTop As Long, _
                             ByVal p_CaptureLast As Long, _
                             ByRef pa_OffsetVector() As Long, _
                             ByVal p_PatternPosition As Long, _
                             ByVal p_NextItemLength As Long, _
                             ByRef p_Action As e_CalloutReceivedAction)
                                   
                               
Private m_Hlib As Long

' Options
Private mo_Options As cPcre2Options

' Cached Handles
Private m_CompiledRegexHandle As Long
Private m_CompileContextHandle As Long
Private m_MatchDataHandle As Long
Private m_MatchContextHandle As Long

Private m_Pattern As String   ' Last used regex pattern

Public Function Options() As cPcre2Options
   If mo_Options Is Nothing Then Set mo_Options = New cPcre2Options
   
   Set Options = mo_Options
End Function

Private Sub Class_Initialize()
   m_Hlib = win32_LoadLibrary("pcre2-16.dll")
   
   If m_Hlib = 0 Then Err.Raise vbObjectError, , "Could not initialize PCRE2 library! Last DLL Error: " & Err.LastDllError
End Sub

Private Sub Class_Terminate()
   Cleanup
   
   If m_Hlib <> 0 Then
      win32_FreeLibrary m_Hlib
   End If
End Sub

Private Function InitRegex(ByVal p_RegexToMatch As String) As Boolean
   Static s_LastRegex As String
   Static s_LastCompileOptionsFlags As Long
   
   ' Returns TRUE if we should continue with match/replace
   ' Returns FALSE if we should abort with no error (host cancelled).
   
   Dim l_ErrorCode As Long
   Dim l_ErrorPosition As Long
   Dim l_Recompile As Boolean
   Dim l_CompileOptionsFlags As Long
   Dim l_CalloutEnumeratedAction As e_CalloutEnumeratedAction
   Dim l_Cancel As Boolean
   Dim l_ObjPtr As Long
   
   ' Use the value of the Pattern property if option p_RegexToMatch parameter is empty
   ' Otherwise set the value of the Pattern property to the passed p_RegexToMatch value
   If p_RegexToMatch = "" Then
      p_RegexToMatch = Me.Pattern
   Else
      Me.Pattern = p_RegexToMatch
   End If
      
   Debug.Print "Regex: " & p_RegexToMatch
   
   l_CompileOptionsFlags = Me.Options.Compile.OptionFlags
   
   ' Check if we should recompile regex
   l_Recompile = True
   If m_CompileContextHandle <> 0 Then
      ' Context exists
      If m_CompiledRegexHandle <> 0 Then
         ' Compile regex exists
         If l_CompileOptionsFlags = s_LastCompileOptionsFlags Then
            If Len(s_LastRegex) = Len(p_RegexToMatch) Then
               If s_LastRegex = p_RegexToMatch Then
                  l_Recompile = False
               End If
            End If
         End If
      End If
   End If
   
   If l_Recompile Then
      'Compile a source string.  Because we are using the 16-bit version, we can just pass StrPtr
      
      Cleanup
      
      m_CompileContextHandle = pcre2_compile_context_create(0)
      If m_CompileContextHandle = 0 Then Err.Raise "Could not compile PCRE context! Last DLL Error: " & Err.LastDllError
      
      m_CompiledRegexHandle = pcre2_compile(StrPtr(p_RegexToMatch), Len(p_RegexToMatch), l_CompileOptionsFlags, l_ErrorCode, l_ErrorPosition, m_CompileContextHandle)
      If m_CompiledRegexHandle = 0 Then Err.Raise vbObjectError, , "Could not compile regex! Regex: " & p_RegexToMatch & vbNewLine & "Errorcode: " & l_ErrorCode & ", Error Position: " & l_ErrorPosition
      
      s_LastCompileOptionsFlags = l_CompileOptionsFlags
      s_LastRegex = p_RegexToMatch
   End If

   ' Create match data space - can we cache and reuse this on subsequent runs?
   If m_MatchDataHandle <> 0 Then pcre2_match_data_free m_MatchDataHandle: m_MatchDataHandle = 0
   
   m_MatchDataHandle = pcre2_match_data_create_from_pattern(m_CompiledRegexHandle, 0)
   
   If m_MatchDataHandle = 0 Then Err.Raise vbObjectError, , "Could not allocate match data! Last DLL Error: " & Err.LastDllError

   l_ObjPtr = ObjPtr(Me)
   
   If Me.Options.General.EnumerateCallouts Then
      ' Enumerate all callouts before matching
      l_CalloutEnumeratedAction = pcre2_callout_enumerate(m_CompiledRegexHandle, AddressOf pcreCalloutEnumerateProc, l_ObjPtr)
      
      Select Case l_CalloutEnumeratedAction
      Case enumcalloutaction_StopAll
         ' Cancel matching.
         l_Cancel = True
         
      Case Is < 0
         ' Raise error
         Err.Raise l_CalloutEnumeratedAction, , "Host reported error."
         
      End Select
   End If
   
   If Me.Options.General.EnableCallouts Then
      ' Callouts enabled during matching/substitution
      
      If m_MatchContextHandle <> 0 Then pcre2_match_context_free m_MatchContextHandle: m_MatchContextHandle = 0
      
      m_MatchContextHandle = pcre2_match_context_create(0)
      
      If m_MatchContextHandle = 0 Then Err.Raise vbObjectError, , "Could not allocate match context space! Last DLL Error: " & Err.LastDllError
      
      Debug.Print "Enabling callouts with ObjPtr: " & l_ObjPtr
      
      pcre2_set_callout m_MatchContextHandle, AddressOf pcreCalloutProc, l_ObjPtr
   End If
   
   InitRegex = Not l_Cancel
End Function

Public Function Replace(ByVal p_TextToSearch As String, ByVal p_ReplaceWithText As String, Optional ByVal p_RegexToMatch As String) As String
   ' Call SUBSTITUTE function instead of this function for new projects.
   ' This function is here for easier migration for existing projects
   ' using VBScript Regex object
   
   Replace = Substitute(p_TextToSearch, p_ReplaceWithText, p_RegexToMatch)
End Function

Public Function Substitute(ByVal p_TextToSearch As String, ByVal p_ReplaceWithText As String, Optional ByVal p_RegexToMatch As String) As String
   ' Calling this function is preferred to calling Replace, even though there is NO
   ' functional difference.
   ' "Substitute" more closely resembles PCRE2 naming,
   ' "Replace" is to allow for easier migration from VBScript Regex object.
   
   ' Returns a string with requested substitutions made (if found).
   ' Make sure GlobalSearch = True to change all matches.
   
   Dim l_ErrorNumber As Long
   Dim l_ErrorDesc As String
   Dim l_ReplaceOptionsFlags As Long
   Dim l_ReplaceCount As Long
   Dim l_OutputBuffer As String
   Dim l_OutputBufferLength As Long
   Dim l_ReplaceResult As Long
   
   On Error GoTo ErrorHandler
   
   Debug.Print "Subject text: " & p_TextToSearch
   
   If Not InitRegex(p_RegexToMatch) Then GoTo NoErrorCleanup

   l_ReplaceOptionsFlags = Me.Options.Replace.OptionFlags
   
   ' Check if we should add the PCRE2_SUBSTITUTE_GLOBAL flag
   If Me.Options.General.GlobalSearch Then l_ReplaceOptionsFlags = l_ReplaceOptionsFlags Or PCRE2_SUBSTITUTE_GLOBAL

   ' Prepare the output buffer (start at 2X size for a better chance to avoid insufficient space)
   l_OutputBuffer = Space$(Len(p_TextToSearch) * 2)
   l_OutputBufferLength = Len(l_OutputBuffer)
   
   ' Attempt substitution
   Do
      l_ReplaceResult = pcre2_substitute(m_CompiledRegexHandle, StrPtr(p_TextToSearch), Len(p_TextToSearch), 0, l_ReplaceOptionsFlags, m_MatchDataHandle, 0, StrPtr(p_ReplaceWithText), Len(p_ReplaceWithText), StrPtr(l_OutputBuffer), l_OutputBufferLength)
      
      Select Case l_ReplaceResult
      Case PCRE_RC_ERROR_NOMEMORY
         ' Buffer too small - increase size.
         
         l_OutputBufferLength = Len(l_OutputBuffer) * 2
         If l_OutputBufferLength > 0 Then
            ' No replacements to make
            l_OutputBuffer = Space$(l_OutputBufferLength)
         End If
         
      
      Case Is >= 0
         ' Finished
         
      Case Else
         Err.Raise vbObjectError - l_ReplaceResult, , "Replace error #" & l_ReplaceResult
         
      End Select
      
   Loop While (l_ReplaceResult = PCRE_RC_ERROR_NOMEMORY) And (l_OutputBufferLength > 0)
   
   Debug.Print "Replacement result: " & l_ReplaceCount   ' NOTE: Documentation says this will return the number of replacements on success, but it always seems to be 0?

   Substitute = Left$(l_OutputBuffer, l_OutputBufferLength)

NoErrorCleanup:
   On Error Resume Next
   
   l_ErrorNumber = 0
   l_ErrorDesc = ""

Cleanup:
   On Error Resume Next

   ' Free match data and context if necessary
   If m_MatchContextHandle <> 0 Then pcre2_match_context_free m_MatchContextHandle: m_MatchContextHandle = 0
   If m_MatchDataHandle <> 0 Then pcre2_match_data_free m_MatchDataHandle: m_MatchDataHandle = 0

   If l_ErrorNumber <> 0 Then
      Cleanup
      
      If IsPcre2ErrorCode(l_ErrorNumber) Then
         l_ErrorDesc = l_ErrorDesc & vbNewLine & "PCRE2 Error Message: " & GetPcre2ErrorMessage(l_ErrorNumber)
      Else
         If IsPcre2ErrorCode(vbObjectError - l_ErrorNumber) Then
            l_ErrorDesc = l_ErrorDesc & vbNewLine & "PCRE2 Error Message: " & GetPcre2ErrorMessage(vbObjectError - l_ErrorNumber)
         End If
      End If
      
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

Public Function Test(ByVal p_TextToSearch As String, Optional ByVal p_RegexToMatch As String) As Boolean
   ' Return TRUE if any match found
   Test = (Execute2(p_TextToSearch, p_RegexToMatch, True).Count > 0)
End Function

Public Function Execute(ByVal p_TextToSearch As String, Optional ByVal p_RegexToMatch As String) As cPcre2Matches
   ' Call MATCH function instead of this function for new projects.
   ' This function is here for easier migration for existing projects
   ' using VBScript Regex object
   
   ' Returns CPcre2Matches object. If MatchCount = 0 then no matches.
   Set Execute = Match(p_TextToSearch, p_RegexToMatch)
End Function

Public Function Match(ByVal p_TextToSearch As String, Optional ByVal p_RegexToMatch As String) As cPcre2Matches
   ' No funcitonal difference between Match and Execute methods, but MATCH is preferred.
   ' "Match" resembles PCRE2 naming, "Execute" resembles VBScript Regex naming
   
   ' Returns CPcre2Matches object. If MatchCount = 0 then no matches.
   Set Match = Execute2(p_TextToSearch, p_RegexToMatch, False)
End Function

Private Function Execute2(ByVal p_TextToSearch As String, Optional ByVal p_RegexToMatch As String, Optional ByVal p_TestOnly As Boolean) As cPcre2Matches
   ' Returns CPcre2Matches object. If MatchCount = 0 then no matches.

   Dim l_ErrorNumber As Long
   Dim l_ErrorDesc As String
   Dim l_MatchCount As Long
   Dim l_OvectorPtr As Long
   Dim la_Ovector() As Long
   Dim l_GlobalSearch As Boolean
   Dim l_Cancel As Boolean
   Dim l_SubstitutionAction As e_SubstitutionAction
   Dim lo_SubstitutionCache As VBA.Collection
   Dim l_Substitute As Boolean
   Dim l_Delta As Long
   Dim ii As Long ' Loop counter
   Dim l_MatchOffset As Long
   Dim l_MatchOptions As Long
   
   On Error GoTo ErrorHandler
   
   Debug.Print "Subject text: " & p_TextToSearch
   
   Set Execute2 = New cPcre2Matches
   
   If Not InitRegex(p_RegexToMatch) Then GoTo NoErrorCleanup
   
   If Not p_TestOnly Then
      l_GlobalSearch = Me.Options.General.GlobalSearch
   End If
   
   l_MatchOptions = Me.Options.Match.OptionFlags   ' Cache match options
   
   Do
      l_MatchCount = pcre2_match(m_CompiledRegexHandle, StrPtr(p_TextToSearch), Len(p_TextToSearch), 0, l_MatchOptions, m_MatchDataHandle, m_MatchContextHandle)
   
      Select Case l_MatchCount
      Case PCRE2_ERROR_NOMATCH
         ' No matches, that's normal :)
      
      Case Is > 0
         ' Number of matches, store information about matches
         l_OvectorPtr = pcre2_get_ovector_pointer(m_MatchDataHandle)
         
         If l_OvectorPtr = 0 Then
            ' Shouldn't happen!
            Err.Raise vbObjectError, , "Ovector pointer could not be retrieved!"
         End If
         
         If Not p_TestOnly Then
            Dim l_SearchChunk As String
            Dim l_MatchStart As Long
            Dim l_MatchLen As Long
            Dim l_MatchText As String
            Dim l_OriginalMatchText As String
            
            win32_CopyMemory l_MatchStart, ByVal l_OvectorPtr, 4
            win32_CopyMemory l_MatchLen, ByVal (l_OvectorPtr + 4), 4
            l_MatchLen = l_MatchLen - l_MatchStart
            
            If l_MatchLen = 0 Then
               ' Empty match
               If Not (Me.Options.Match.MatchIfEmpty Or Me.Options.Match.MatchIfEmptyAtStart) Then
                  ' We don't want empty matches, so go to cleanup
                  GoTo NoErrorCleanup
               End If
            End If
            
            l_SearchChunk = Left$(p_TextToSearch, l_MatchStart + l_MatchLen)
            l_MatchText = Mid$(l_SearchChunk, l_MatchStart + 1)
            l_OriginalMatchText = l_MatchText
            
            If Me.Options.Match.MatchedEventEnabled Then
               ' Check if we have a substitution stored in our cache
               l_Substitute = False
               
               If lo_SubstitutionCache Is Nothing Then
                  Set lo_SubstitutionCache = New VBA.Collection
                  l_ErrorNumber = 5 ' Simulate not found in collection error
               Else
                  On Error Resume Next
                  Err.Clear
                  l_MatchText = lo_SubstitutionCache.Item("KEY_" & l_OriginalMatchText)
                  l_ErrorNumber = Err.Number
                  On Error GoTo ErrorHandler
               End If
               
               If l_ErrorNumber = 0 Then
                  Debug.Print "Substitution cache hit for " & l_OriginalMatchText & ". Skipping Matched event and using cached value."
                  
                  l_Substitute = True
               Else
                  ' Could not find substitution match in cache
                  ' Raise the Matched event since we found a match
                  
                  l_ErrorNumber = 0
                  
                  l_SubstitutionAction = subaction_None
                  RaiseEvent Matched(l_MatchText, l_SubstitutionAction, l_Cancel)
                  
                  If l_Cancel Then GoTo NoErrorCleanup   ' Received instruction to cancel matching
                  
                  Select Case l_SubstitutionAction
                  Case subaction_ReplaceAndCache, subaction_Replace
                     If l_SubstitutionAction = subaction_ReplaceAndCache Then
                        lo_SubstitutionCache.Add l_MatchText, "KEY_" & l_OriginalMatchText
                           
                        ' Make sure cache doesn't get too big
                        For ii = Me.Options.Match.MatchedEventSubstitutionCacheSize - 1 To lo_SubstitutionCache.Count - 1
                           lo_SubstitutionCache.Remove 0
                        Next ii
                     
                     End If
                     
                     l_Substitute = True
                     
                  Case subaction_None
                     ' Make sure match text is unchanged.
                     l_MatchText = l_OriginalMatchText
                  
                  Case Else
                     Err.Raise vbObjectError, , "Unknown substitution action: " & l_MatchText
                  End Select
               End If
            End If
         End If
         
         ReDim la_Ovector(2 * l_MatchCount - 1)

         win32_CopyMemory la_Ovector(0), ByVal l_OvectorPtr, 2 * l_MatchCount * 4
         
         If l_Substitute Then
            l_Delta = Len(l_MatchText) - l_MatchLen
            
            If l_Delta > 0 Then
               p_TextToSearch = p_TextToSearch & Space$(l_Delta)
               
               Mid$(p_TextToSearch, l_MatchStart + l_MatchLen + l_Delta + 1) = Mid$(p_TextToSearch, l_MatchStart + l_MatchLen + 1)
               
            End If
            
            Mid$(p_TextToSearch, l_MatchStart + 1) = l_MatchText
         
            If l_Delta < 0 Then
               Mid$(p_TextToSearch, l_MatchStart + l_MatchLen + l_Delta + 1) = Mid$(p_TextToSearch, l_MatchStart + l_MatchLen + 1)
               p_TextToSearch = Left$(p_TextToSearch, Len(p_TextToSearch) + l_Delta)
            End If
         
         Else
            Execute2.Add l_SearchChunk, la_Ovector, l_MatchCount, l_MatchOffset
            
            l_MatchOffset = l_MatchOffset + l_MatchStart + l_MatchLen
         
            If l_GlobalSearch Then
               p_TextToSearch = Mid$(p_TextToSearch, l_MatchStart + l_MatchLen + 1)
            End If
         
         End If
         
      Case Else
         ' Uhoh! We need to handle these
         Err.Raise vbObjectError - l_MatchCount, , "PCRE Match Error: " & l_MatchCount
      End Select
      
   ' Loop only if global search enabled and we have a match and a match length.
   Loop While l_GlobalSearch And (l_MatchCount > 0) And (l_MatchLen > 0)
   
NoErrorCleanup:
   On Error Resume Next
   
   Execute2.AppendText p_TextToSearch
   
   l_ErrorNumber = 0
   l_ErrorDesc = ""
   
Cleanup:
   On Error Resume Next

   ' Free match data if necessary
   If m_MatchContextHandle <> 0 Then pcre2_match_context_free m_MatchContextHandle: m_MatchContextHandle = 0
   If m_MatchDataHandle <> 0 Then pcre2_match_data_free m_MatchDataHandle: m_MatchDataHandle = 0

   If l_ErrorNumber <> 0 Then
      Cleanup
      
      If IsPcre2ErrorCode(l_ErrorNumber) Then
         l_ErrorDesc = l_ErrorDesc & vbNewLine & "PCRE2 Error Message: " & GetPcre2ErrorMessage(l_ErrorNumber)
      Else
         If IsPcre2ErrorCode(vbObjectError - l_ErrorNumber) Then
            l_ErrorDesc = l_ErrorDesc & vbNewLine & "PCRE2 Error Message: " & GetPcre2ErrorMessage(vbObjectError - l_ErrorNumber)
         End If
      End If
      
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

Private Function IsPcre2ErrorCode(ByVal p_ErrorCode As Long) As Boolean
   IsPcre2ErrorCode = (p_ErrorCode <= [_PCRE_RC_ERROR_FIRST] And p_ErrorCode >= [_PCRE_RC_ERROR_LAST])
End Function

Private Function GetPcre2ErrorMessage(ByVal p_ErrorCode As Long) As String
   Dim l_BufferLength As Long
   Dim l_Buffer As String
   Dim l_MessageLength As Long
   
   l_BufferLength = 256
   
   Do
      l_Buffer = Space$(l_BufferLength)
      
      l_MessageLength = pcre2_get_error_message(p_ErrorCode, StrPtr(l_Buffer), l_BufferLength)
      
      If l_MessageLength < 0 Then
         Select Case l_MessageLength
         Case PCRE_RC_ERROR_NOMEMORY
            ' Buffer too small
            l_BufferLength = l_BufferLength * 2
         Case PCRE_RC_ERROR_BADDATA
            ' Bad error code
            
            Exit Do
         Case Else
            Debug.Assert False
            Exit Do
            
         End Select
      End If
   Loop While l_MessageLength < 0
   
   If l_MessageLength < 0 Then
      GetPcre2ErrorMessage = "Unknown error #" & p_ErrorCode & ", PCRE2 error message result #" & l_MessageLength
   Else
      GetPcre2ErrorMessage = Left$(l_Buffer, l_MessageLength)
   End If
End Function

Private Sub Cleanup()
   'Free compiled regex before exiting
   If m_CompiledRegexHandle <> 0 Then pcre2_code_free m_CompiledRegexHandle: m_CompiledRegexHandle = 0

   'Free compile context before exiting
   If m_CompileContextHandle <> 0 Then pcre2_compile_context_free m_CompileContextHandle: m_CompileContextHandle = 0
End Sub

Friend Function RaiseCalloutEnumeratedEvent(pt_CalloutEnumerateBlock As pcreCalloutEnumerateBlock) As e_CalloutEnumeratedAction
   Dim l_CalloutLabel As String
   Dim l_Action As e_CalloutEnumeratedAction
   
   With pt_CalloutEnumerateBlock
      If .CalloutStringPointer <> 0 Then
         l_CalloutLabel = Left$(stringGetFromPointerW(.CalloutStringPointer), .CalloutStringLength)
      End If
      
      RaiseEvent CalloutEnumerated(.CalloutNumber, l_CalloutLabel, .CalloutStringOffset, .PatternPosition, .NextItemLength, l_Action)
   End With
   
   RaiseCalloutEnumeratedEvent = l_Action
End Function

Friend Function RaiseCalloutReceivedEvent(pt_CalloutBlock As pcreCalloutBlock) As e_CalloutReceivedAction
   Dim l_CalloutLabel As String
   Dim l_Mark As String
   Dim l_Subject As String
   Dim l_Action As e_CalloutReceivedAction
   Dim l_OvectorCount As Long
   Dim la_OffsetVector() As Long
   
   With pt_CalloutBlock
      If .CalloutStringPointer <> 0 Then
         If .CalloutStringLength > 0 Then
            l_CalloutLabel = Left$(stringGetFromPointerW(.CalloutStringPointer), .CalloutStringLength)
         End If
      End If
      
      If .SubjectPointer <> 0 Then
         If .SubjectLength > 0 Then
            l_Subject = Left$(stringGetFromPointerW(.SubjectPointer), .SubjectLength)
         End If
      End If
      
      If .MarkPointer <> 0 Then
         l_Mark = stringGetFromPointerW(.MarkPointer)
      End If
      
      If m_MatchDataHandle <> 0 Then
         l_OvectorCount = pcre2_get_ovector_count(m_MatchDataHandle)
         
         ReDim la_Ovector(2 * l_OvectorCount - 1)
   
         win32_CopyMemory la_Ovector(0), ByVal .OffsetVectorPointer, 2 * l_OvectorCount * 4
      Else
         ReDim la_Ovector(0 To -1)
      End If
      
      RaiseEvent CalloutReceived(.CalloutNumber, l_CalloutLabel, .CalloutStringOffset, l_Subject, l_Mark, .CaptureTop, .CaptureLast, la_OffsetVector, .PatternPosition, .NextItemLength, l_Action)
   End With
   
   RaiseCalloutReceivedEvent = l_Action
End Function

' GlobalSearch and IgnoreCase properties are aliases for GlobalSearch and CaseSensitive options
' To more closely match VBScript Regex Object Model. Unfortunately we can't have a property call Global
' because it is a reserved keyword, so we can match VBScript Regex OM exactly :(
Public Property Let GlobalSearch(ByVal p_Enable As Boolean)
   Me.Options.General.GlobalSearch = p_Enable
End Property

Public Property Get GlobalSearch() As Boolean
   GlobalSearch = Me.Options.General.GlobalSearch
End Property

Public Property Let IgnoreCase(ByVal p_Enable As Boolean)
   Me.Options.Compile.CaseSensitive = Not p_Enable
End Property

Public Property Get IgnoreCase() As Boolean
   IgnoreCase = Not Me.Options.Compile.CaseSensitive
End Property

' Pattern property added to more closely mimic VBScript Regex object model
Public Property Get Pattern() As String
   Pattern = m_Pattern
End Property

Public Property Let Pattern(ByVal p_RegexPattern As String)
   m_Pattern = p_RegexPattern
End Property


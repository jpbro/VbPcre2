VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPcre2OptionsMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

' Some documentation from http://www.pcre.org/current/doc/html/pcre2api.html
' Some additional documentation for PCRE2 Regex options by Seongmyun Cho taken from https://github.com/xnsystems/kpcre/wiki

Private m_Options As Long
Private m_MatchedEventEnabled As Boolean
Private m_SubstitutionCacheSize As Long

Private Sub Class_Initialize()
   ' Setup default values
   Me.MatchedEventSubstitutionCacheSize = 100
   Me.SubjectIsBeginningOfLine = True
   Me.SubjectIsEndOfLine = True
   Me.MatchIfEmpty = True
   Me.MatchIfEmptyAtStart = True
   Me.PerformUtfCheck = True
End Sub

Friend Function OptionFlags() As Long
   OptionFlags = m_Options
End Function

Private Sub AddOrRemoveRegexOptionFlag(ByVal p_Flag As Long, ByVal p_Add As Boolean)
   If p_Add Then
      m_Options = m_Options Or p_Flag
   Else
      m_Options = m_Options And Not p_Flag
   End If
End Sub

' If ANCHORED option is TRUE, the pattern is forced to be "anchored", that is, it is constrained to match only
' at the first matching point in the string that is being searched.
Public Property Get Anchored() As Boolean
   Anchored = m_Options And PCRE2_ANCHORED
End Property

Public Property Let Anchored(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_ANCHORED, p_Enabled
End Property


Public Property Get SubjectIsBeginningOfLine() As Boolean
   SubjectIsBeginningOfLine = Not CBool(m_Options And PCRE2_NOTBOL)
End Property

Public Property Let SubjectIsBeginningOfLine(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_NOTBOL, Not p_Enabled
End Property

Public Property Get SubjectIsEndOfLine() As Boolean
   SubjectIsEndOfLine = Not CBool(m_Options And PCRE2_NOTEOL)
End Property

Public Property Let SubjectIsEndOfLine(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_NOTEOL, Not p_Enabled
End Property

Public Property Get MatchIfEmpty() As Boolean
   MatchIfEmpty = Not CBool(m_Options And PCRE2_NOTEMPTY)
End Property

Public Property Let MatchIfEmpty(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_NOTEMPTY, Not p_Enabled
End Property

Public Property Get MatchIfEmptyAtStart() As Boolean
   MatchIfEmptyAtStart = Not CBool(m_Options And PCRE2_NOTEMPTY_ATSTART)
End Property

Public Property Let MatchIfEmptyAtStart(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_NOTEMPTY_ATSTART, Not p_Enabled
End Property

Public Property Get PerformUtfCheck() As Boolean
   PerformUtfCheck = Not CBool(m_Options And PCRE2_NO_UTF_CHECK)
End Property

Public Property Let PerformUtfCheck(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_NO_UTF_CHECK, Not p_Enabled
End Property

' Set MatchedEventEnabled to True to enable the Matched() event to be raised after each match is found
' NOTE: This only makes sense to use when the Global property is also set to True.
' Set MatchedEventEnabled to False to skip raising the Matched event - should improve performance if you don't need the event.
Public Property Get MatchedEventEnabled() As Boolean
   MatchedEventEnabled = m_MatchedEventEnabled
End Property
   
Public Property Let MatchedEventEnabled(ByVal p_Enabled As Boolean)
   m_MatchedEventEnabled = p_Enabled
End Property
   

Public Property Get MatchedEventSubstitutionCacheSize() As Long
   MatchedEventSubstitutionCacheSize = m_SubstitutionCacheSize
End Property

Public Property Let MatchedEventSubstitutionCacheSize(ByVal p_MaxCount As Long)
   m_SubstitutionCacheSize = p_MaxCount
End Property



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPcreOptionsReplace"
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
' Some additional documentation from http://manpages.ubuntu.com/manpages/zesty/man3/pcre2api.3.html

Private Const PCRE2_SUBSTITUTE_EXTENDED As Long = &H200
Private Const PCRE2_SUBSTITUTE_UNSET_EMPTY As Long = &H400
Private Const PCRE2_SUBSTITUTE_UNKNOWN_UNSET As Long = &H800
' NOT NEEDED - We handle buffer overflows transparently for host. Private Const PCRE2_SUBSTITUTE_OVERFLOW_LENGTH As Long = &H1000

Private m_Options As Long

Private Sub Class_Initialize()
   ' Setup default values
   Me.SubjectIsBeginningOfLine = True
   Me.SubjectIsEndOfLine = True
   Me.MatchIfEmpty = True
   Me.MatchIfEmptyAtStart = True
   Me.PerformUtfCheck = True

   Me.ErrorOnUnknownCapturingGroups = True
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


' PCRE2_SUBSTITUTE_UNKNOWN_UNSET  causes  references  to capturing groups
' that do not appear in the pattern to be treated as unset  groups.  This
' option  should  be  used  with  care, because it means that a typo in a
' group name or  number  no  longer  causes  the  PCRE2_ERROR_NOSUBSTRING
' error.
Public Property Get ErrorOnUnknownCapturingGroups() As Boolean
   ErrorOnUnknownCapturingGroups = Not CBool(m_Options And PCRE2_SUBSTITUTE_UNKNOWN_UNSET)
End Property

Public Property Let ErrorOnUnknownCapturingGroups(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_SUBSTITUTE_UNKNOWN_UNSET, Not p_Enabled
End Property

' PCRE2_SUBSTITUTE_UNSET_EMPTY  causes  unset capturing groups (including
' unknown  groups  when  PCRE2_SUBSTITUTE_UNKNOWN_UNSET  is  set)  to  be
' treated  as  empty  strings  when  inserted as described above. If this
' option is not set, an attempt to  insert  an  unset  group  causes  the
' PCRE2_ERROR_UNSET  error.  This  option does not influence the extended
' substitution syntax.
Public Property Get TreatUnknownCapturingGroupsAsEmptyStrings() As Boolean
   TreatUnknownCapturingGroupsAsEmptyStrings = m_Options And PCRE2_SUBSTITUTE_UNSET_EMPTY
End Property

Public Property Let TreatUnknownCapturingGroupsAsEmptyStrings(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_SUBSTITUTE_UNSET_EMPTY, p_Enabled
End Property


' PCRE2_SUBSTITUTE_EXTENDED causes extra processing to be applied to  the
' replacement  string.  Without this option, only the dollar character is
' special, and only the group insertion forms  listed  above  are  valid.
' When PCRE2_SUBSTITUTE_EXTENDED is set, two things change:
'
'       Firstly,  backslash in a replacement string is interpreted as an escape
'       character. The usual forms such as \n or \x{ddd} can be used to specify
'       particular   character  codes,  and  backslash  followed  by  any  non-
'       alphanumeric character quotes that character. Extended quoting  can  be
'       coded using \Q...\E, exactly as in pattern strings.
'
'       There  are  also four escape sequences for forcing the case of inserted
'       letters.  The insertion mechanism has three states:  no  case  forcing,
'       force upper case, and force lower case. The escape sequences change the
'       current state: \U and  \L  change  to  upper  or  lower  case  forcing,
'       respectively,  and  \E  (when  not  terminating  a  \Q quoted sequence)
'       reverts to no case forcing. The sequences \u  and  \l  force  the  next
'       character (if it is a letter) to upper or lower case, respectively, and
'       then the state automatically reverts to no case forcing.  Case  forcing
'       applies  to  all  inserted   characters,  including those from captured
'       groups and letters within \Q...\E quoted sequences.
'
'       Note that case forcing sequences such  as  \U...\E  do  not  nest.  For
'       example,  the  result  of  processing "\Uaa\LBB\Ecc\E" is "AAbbcc"; the
'       final \E has no effect.
'
'       The second effect of setting PCRE2_SUBSTITUTE_EXTENDED is to  add  more
'       flexibility  to  group substitution. The syntax is similar to that used
'       By Bash:
'
'         ${<n>:-<string>}
'         ${<n>:+<string1>:<string2>}
'
'       As before, <n> may be  a  group  number  or  a  name.  The  first  form
'       specifies  a default value. If group <n> is set, its value is inserted;
'       if not, <string> is expanded and the result inserted. The  second  form
'       specifies  strings that are expanded and inserted when group <n> is set
'       or unset, respectively. The first form is just a  convenient  shorthand
'       for
'
'         ${<n>:+${<n>}:<string>}
'
'       Backslash  can  be  used to escape colons and closing curly brackets in
'       the replacement strings. A change of the case forcing  state  within  a
'       replacement  string  remains  in  force  afterwards,  as  shown in this
'       pcre2test example:
'
'         /(some)?(body)/substitute_extended,replace=${1:+\U:\L}HeLLo
'             body
'             1:HELLO
'             somebody
'             1:HELLO
'
'       The PCRE2_SUBSTITUTE_UNSET_EMPTY option does not affect these  extended
'       substitutions.   However,   PCRE2_SUBSTITUTE_UNKNOWN_UNSET  does  cause
'       unknown groups in the extended syntax forms to be treated as unset.

Public Property Get ExtendedReplacement() As Boolean
   ExtendedReplacement = m_Options And PCRE2_SUBSTITUTE_EXTENDED
End Property

Public Property Let ExtendedReplacement(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_SUBSTITUTE_EXTENDED, p_Enabled
End Property



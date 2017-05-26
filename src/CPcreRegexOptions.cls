VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPcreRegexOptions"
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

' Some documentation from http://www.pcre.org/current/doc/html/pcre2api.html
' Some additional documentation for PCRE2 Regex options by Seongmyun Cho taken from https://github.com/xnsystems/kpcre/wiki

Private Const PCRE2_ANCHORED            As Long = &H80000000
Private Const PCRE2_NO_UTF_CHECK        As Long = &H40000000
Private Const PCRE2_ALLOW_EMPTY_CLASS   As Long = &H1
Private Const PCRE2_ALT_BSUX            As Long = &H2
Private Const PCRE2_CASELESS            As Long = &H8
Private Const PCRE2_DOLLAR_ENDONLY      As Long = &H10
Private Const PCRE2_DOTALL              As Long = &H20
Private Const PCRE2_EXTENDED            As Long = &H80
Private Const PCRE2_FIRSTLINE           As Long = &H100
Private Const PCRE2_MULTILINE           As Long = &H400
Private Const PCRE2_UNGREEDY            As Long = &H40000
Private Const PCRE2_UTF                 As Long = &H80000
Private Const PCRE2_NEVER_BACKSLASH_C   As Long = &H100000
Private Const PCRE2_ALT_CIRCUMFLEX      As Long = &H200000
Private Const PCRE2_AUTO_CALLOUT        As Long = &H4 ' TODO: Callouts not implemented

' UNIMPLEMENTED
Private Const PCRE2_DUPNAMES            As Long = &H40
Private Const PCRE2_MATCH_UNSET_BACKREF As Long = &H200
Private Const PCRE2_NEVER_UCP           As Long = &H800
Private Const PCRE2_NEVER_UTF           As Long = &H1000
Private Const PCRE2_NO_AUTO_CAPTURE     As Long = &H2000
Private Const PCRE2_NO_AUTO_POSSESS     As Long = &H4000
Private Const PCRE2_NO_DOTSTAR_ANCHOR   As Long = &H8000
Private Const PCRE2_NO_START_OPTIMIZE   As Long = &H10000
Private Const PCRE2_UCP                 As Long = &H20000

Private m_Options As Long
Private m_Global As Boolean

Private Sub Class_Initialize()
   ' Setup default options
   
   m_Options = PCRE2_NEVER_BACKSLASH_C ' Recommended always on for UTF
   
   Me.Utf = True
   Me.CheckUtfValidity = True
   Me.CaseSensitive = True
   Me.Greedy = True
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

' If CASE SENSITIVE option is FALSE, letters in the pattern match both upper and lower case letters in the subject.
Public Property Get CaseSensitive() As Boolean
   CaseSensitive = m_Options And Not PCRE2_CASELESS
End Property

Public Property Let CaseSensitive(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_CASELESS, Not p_Enabled
End Property


' If ANCHORED option is TRUE, the pattern is forced to be "anchored", that is, it is constrained to match only
' at the first matching point in the string that is being searched.
Public Property Get Anchored() As Boolean
   Anchored = m_Options And PCRE2_ANCHORED
End Property

Public Property Let Anchored(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_ANCHORED, p_Enabled
End Property


' If the PCRE2_AUTO_CALLOUT option bit is set when a pattern is compiled, PCRE2
' automatically inserts callouts, all with number 255, before each item in the
' pattern. For example, if PCRE2_AUTO_CALLOUT is used with the pattern
'
'   A(\ed{2}|--)
'
' it is processed as if it were
'
' (?C255)A(?C255)((?C255)\ed{2}(?C255)|(?C255)-(?C255)-(?C255))(?C255)
'
' Notice that there is a callout before and after each parenthesis and
' alternation bar. If the pattern contains a conditional group whose condition is
' an assertion, an automatic callout is inserted immediately before the
' condition. Such a callout may also be inserted explicitly, for example:
'
'   (?(?C9)(?=a)ab|de)
'
' This applies only to assertion conditions (because they are themselves
' independent groups).
'
' Automatic callouts can be used for tracking the progress of pattern matching.
Public Property Get AutomaticallyInsertCallouts() As Boolean
   AutomaticallyInsertCallouts = m_Options And PCRE2_AUTO_CALLOUT
End Property

Public Property Let AutomaticallyInsertCallouts(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_AUTO_CALLOUT, p_Enabled
End Property


' If the IgnorePatternWhitspaceAndComments option is TRUE, most white space characters in the pattern are totally ignored except when escaped or
' inside a character class. However, white space is not allowed within sequences such as (?> that introduce various
' parenthesized subpatterns, nor within numerical quantifiers such as {1,3}.
' Ignorable white space is permitted between an item and a following quantifier and between a quantifier and
' a following + that indicates possessiveness.
' This setting also causes characters between an unescaped # outside a character class and the next newline,
' inclusive, to be ignored, which makes it possible to include comments inside complicated patterns.
' Note that the end of this type of comment is a literal newline sequence in the pattern; escape sequences that happen to represent a newline do not count.
' PCRE2_EXTENDED is equivalent to Perl's /x option, and it can be changed within a pattern by a (?x) option setting.
' Which characters are interpreted as newlines can be specified by a setting in the compile context that is passed to pcre2_compile() or by a special sequence at the start of the pattern, as described in the section entitled "Newline conventions" in the pcre2pattern documentation.
' A default is defined when PCRE2 is built.
Public Property Get IgnorePatternWhitspaceAndComments() As Boolean
   IgnorePatternWhitspaceAndComments = m_Options And PCRE2_EXTENDED
End Property

Public Property Let IgnorePatternWhitspaceAndComments(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_EXTENDED, p_Enabled
End Property


' If the DOTALL option is TRUE, a dot metacharacter in the pattern matches any character, including one that indicates
' a newline. However, it only ever matches one character, even if newlines are coded as CRLF.
' Without this option, a dot does not match when the current position in the subject is at a newline.
' This option is equivalent to Perl's /s option, and it can be changed within a pattern by a (?s) option setting.
' A negative class such as [^a] always matches newline characters, independent of the setting of this option.
Public Property Get DotMatchesAllCharacters() As Boolean
   DotMatchesAllCharacters = m_Options And PCRE2_DOTALL
End Property

Public Property Let DotMatchesAllCharacters(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_DOTALL, p_Enabled
End Property


' This option inverts the "greediness" of the quantifiers so that they are not greedy by default,
' but become greedy if followed by "?".
Public Property Get Greedy() As Boolean
   Greedy = m_Options And Not PCRE2_UNGREEDY
End Property

Public Property Let Greedy(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_UNGREEDY, Not p_Enabled
End Property


' When PCRE2_UTF is set, the validity of the pattern as a UTF string is automatically checked.
' There are discussions about the validity of UTF-8 strings, UTF-16 strings, and UTF-32 strings in the pcre2unicode document.
' If an invalid UTF sequence is found, pcre2_compile() returns a negative error code.
' If you know that your pattern is valid, and you want to skip this check for performance reasons, you can set the PCRE2_NO_UTF_CHECK option.
' When it is set, the effect of passing an invalid UTF string as a pattern is undefined.
' It may cause your program to crash or loop. Note that this option can also be passed to pcre2_match() and pcre_dfa_match(), to suppress validity checking of the subject string.
Public Property Get CheckUtfValidity() As Boolean
   CheckUtfValidity = m_Options And Not PCRE2_NO_UTF_CHECK
End Property

Public Property Let CheckUtfValidity(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_NO_UTF_CHECK, Not p_Enabled
End Property


' By default, for the purposes of matching "start of line" and "end of line", PCRE2 treats the subject string as
' consisting of a single line of characters, even if it actually contains newlines. The "start of line" metacharacter
' (^) matches only at the start of the string, and the "end of line" metacharacter ($) matches only at the end of the
' string, or before a terminating newline (except when PCRE2_DOLLAR_ENDONLY is set).
' Note, however, that unless PCRE2_DOTALL is set, the "any character" metacharacter (.) does not match at a newline.
Public Property Get Multiline() As Boolean
   Multiline = m_Options And PCRE2_MULTILINE
End Property

Public Property Let Multiline(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_MULTILINE, p_Enabled
End Property


' By default, for compatibility with Perl, a closing square bracket that immediately follows an opening one
' is treated as a data character for the class. When PCRE2_ALLOW_EMPTY_CLASS is set, it terminates the class,
' which therefore contains no characters and so can never match.
Public Property Get AllowEmptyClass() As Boolean
   AllowEmptyClass = m_Options And PCRE2_ALLOW_EMPTY_CLASS
End Property

Public Property Let AllowEmptyClass(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_ALLOW_EMPTY_CLASS, p_Enabled
End Property


' If the AlternateBSUXHandling option is TRUE alternative handling of three escape sequences will be enabled.
' This makes PCRE2's behaviour more like ECMAscript (aka JavaScript). When it is set:
'(1) \U matches an upper case "U" character; by default \U causes a compile time error (Perl uses \U to upper case subsequent characters).
'
'(2) \u matches a lower case "u" character unless it is followed by four hexadecimal digits,
'    in which case the hexadecimal number defines the code point to match.
'    By default, \u causes a compile time error (Perl uses it to upper case the following character).
'
'(3) \x matches a lower case "x" character unless it is followed by two hexadecimal digits,
'    in which case the hexadecimal number defines the code point to match.
'    By default, as in Perl, a hexadecimal number is always expected after \x, but it may have zero, one, or two digits (so, for example, \xz matches a binary zero character followed by z).
Public Property Get AlternateBsuxHandling() As Boolean
   AlternateBsuxHandling = m_Options And PCRE2_ALT_BSUX
End Property

Public Property Let AlternateBsuxHandling(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_ALT_BSUX, p_Enabled
End Property


' If the DOLLAR_ENDONLY option is set, a dollar metacharacter in the pattern matches only at the end of the subject string.
' Without this option, a dollar also matches immediately before a newline at the end of the string
' ( but not before any other newlines). The PCRE2_DOLLAR_ENDONLY option is ignored if PCRE2_DollarEndOnly is set.
Public Property Get DollarMatchesEndOfStringOnly() As Boolean
   DollarMatchesEndOfStringOnly = m_Options And PCRE2_DOLLAR_ENDONLY
End Property

Public Property Let DollarMatchesEndOfStringOnly(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_DOLLAR_ENDONLY, p_Enabled
End Property


' This option causes PCRE2 to regard both the pattern and the subject strings that are subsequently processed as strings of UTF characters instead of single-code-unit strings.
' It is available when PCRE2 is built to include Unicode support (which is the default).
' If Unicode support is not available, the use of this option provokes an error.
' Details of how this option changes the behaviour of PCRE2 are given in the pcre2unicode page.
Public Property Get Utf() As Boolean
   Utf = m_Options And PCRE2_UTF
End Property

Public Property Let Utf(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_UTF, p_Enabled
End Property


' In multiline mode (when PCRE2_MULTILINE is set), the circumflex metacharacter matches at the start of the subject
' (unless PCRE2_NOTBOL is set), and also after any internal newline.
' However, it does not match after a newline at the end of the subject, for compatibility with Perl.
' If you want a multiline circumflex also to match after a terminating newline, you must set PCRE2_ALT_CIRCUMFLEX.
Public Property Get AlternateCircumflexHandling() As Boolean
   AlternateCircumflexHandling = m_Options And PCRE2_ALT_CIRCUMFLEX
End Property

Public Property Let AlternateCircumflexHandling(ByVal p_Enabled As Boolean)
   AddOrRemoveRegexOptionFlag PCRE2_ALT_CIRCUMFLEX, p_Enabled
End Property


' Set GlobalSearch to True to find all possible matches
' Set GlobalSearch to False to stop search after first match.
Public Property Get GlobalSearch() As Boolean
   GlobalSearch = m_Global
End Property
   
Public Property Let GlobalSearch(ByVal p_Enabled As Boolean)
   m_Global = p_Enabled
End Property
   

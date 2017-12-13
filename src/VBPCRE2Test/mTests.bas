Attribute VB_Name = "mTests"
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

Sub TestRegexMatchMultiline()
   Dim lo_RegEx As New cPcre2
   
   lo_RegEx.Options.Compile.MultiLine = True
   
   Debug.Print lo_RegEx.match("XXXAXXXX" & vbNewLine & "YYYYYAYYYYY", "A.+A").Count
End Sub

Sub TestRegexReplace()
   Dim lo_RegEx As New cPcre2
   Dim lo_Matches As cPcre2Matches
   
   With lo_RegEx.Options.Compile
      .CaseSensitive = False
   End With
   
   lo_RegEx.GlobalSearch = False
   
   Set lo_Matches = lo_RegEx.match("This is a test.", "t")
   Debug.Print lo_Matches.Text
   
   Debug.Print "Replace result: " & lo_RegEx.Substitute(lo_Matches.Text, "X")
End Sub

Sub TestRegexMatch()
   Dim lo_RegEx As New cPcre2
   Dim lo_Matches As cPcre2Matches
   Dim ii As Long

   With lo_RegEx.Options.Compile
      .CaseSensitive = False
   End With
   
   Set lo_Matches = lo_RegEx.Execute("This is a test of matching stuff!", "(test)\s*.+\s*(Mat)")
   
   Debug.Print lo_Matches.Item(0).SubMatchFirstIndex(0)
   
   If lo_Matches.Count > 0 Then
      For ii = 0 To lo_Matches.Count - 1
         Debug.Print "Match #" & ii + 1 & ": " & lo_Matches(ii).MatchedText
      Next ii
      
   Else
      Debug.Print "No matches found!"
   End If
End Sub

Sub TestRegexCallout(po_Pcre As cPcre2)
   Dim lo_Matches As cPcre2Matches
   Dim ii As Long
   
   With po_Pcre.Options.Compile
      .CaseSensitive = False
   End With
   
   With po_Pcre.Options.General
      .EnableCallouts = True
   End With
   
   Set lo_Matches = po_Pcre.Execute("This is a test of matching stuff!", "(?C""test"")test\s*.+\s*(Mat)")
   If lo_Matches.Count > 0 Then
      For ii = 0 To lo_Matches.Count - 1
         Debug.Print "Match #" & ii + 1 & ": " & lo_Matches(ii).MatchedText
      Next ii
      
   Else
      Debug.Print "No matches found!"
   End If
End Sub

Sub TestRegexEnumerateCallout(po_Pcre As cPcre2)
   Dim lo_Matches As cPcre2Matches
   Dim ii As Long
   
   With po_Pcre.Options.Compile
      .CaseSensitive = False
   End With
   
   With po_Pcre.Options.General
      .EnumerateCallouts = True
   End With
   
   Set lo_Matches = po_Pcre.Execute("This is a test of matching stuff!", "(?C""test"")test\s*.+\s*(Mat)")
   If lo_Matches.Count > 0 Then
      For ii = 0 To lo_Matches.Count - 1
         Debug.Print "Match #" & ii + 1 & ": " & lo_Matches(ii).MatchedText
      Next ii
      
   Else
      Debug.Print "No matches found!"
   End If
End Sub

Sub TestRegexMatchedEvent(po_Pcre As cPcre2)
   Dim lo_Matches As cPcre2Matches
   
   With po_Pcre.Options.Compile
      .CaseSensitive = False
   End With
   
   With po_Pcre.Options.General
      .GlobalSearch = True
   End With
   
   With po_Pcre.Options.match
      .MatchedEventEnabled = True
   End With
   
   Set lo_Matches = po_Pcre.Execute("This is a test of CB9CBC213211BF4BA026FED4B1AC5CB2. Did you know that CB9CBC213211BF4BA026FED4B1AC5CB2 can be cached for performance, or re-run for each match. F8CEC6F354497746990BFB3A6A72BD06, F8CEC6F354497746990BFB3A6A72BD06, F8CEC6F354497746990BFB3A6A72BD06. You can also ignore matches: E175FA8438E00D47B2AA52CAD413FB6A.", "[a-fA-F0-9]{32}")

   Debug.Print "Result text: " & lo_Matches.Text
End Sub

Sub TestRegex2()
   Dim lo_RegEx As Object ' VBScript_RegExp_55.RegExp
   Dim lo_Matches As Object 'VBScript_RegExp_55.MatchCollection
   Dim lo_Match As Object 'VBScript_RegExp_55.Match
   
   Dim lo_RegEx2 As cPcre2
   Dim lo_Matches2 As cPcre2Matches
   Dim lo_Match2 As cPcre2Match

   Dim l_SubjectText As String
   Dim l_Regex As String
   
   Dim ii As Long
   Dim jj As Long
   
   l_SubjectText = "Wyzo" '"File1.zip.exe" & vbCrLf & "File2.com" & vbCrLf & "File 3"
   l_Regex = "(Manager)|^((copy of )?(" & "MX5|Wyzo" & "))$" '".*$"  '"[\w ]+(\.\S+?)*$"
   
   l_SubjectText = "File1.zip.exe" & vbCrLf & "File2.com" & vbCrLf & "File 3"
   l_Regex = ".*$"  '"[\w ]+(\.\S+?)*$"
   
   ' VBScript Test
   Debug.Print "VBSCRIPT Test"
   
   Set lo_RegEx = CreateObject("VBScript.RegExp")
   With lo_RegEx
      .IgnoreCase = True
      .Global = True
      .MultiLine = True
   End With
   
   lo_RegEx.Pattern = l_Regex
   
   Set lo_Matches = lo_RegEx.Execute(l_SubjectText)
   
   Debug.Print "Match Count: " & lo_Matches.Count
         
   For ii = 0 To lo_Matches.Count - 1
      Set lo_Match = lo_Matches(ii)
      
      Debug.Print "Match #" & ii + 1 & ": " & lo_Match.Value
      Debug.Print "Sub Match Count: " & lo_Match.SubMatches.Count
      For jj = 0 To lo_Match.SubMatches.Count - 1
         Debug.Print "SubMatch # " & jj + 1 & ": " & lo_Match.SubMatches(jj)
      Next jj
   Next ii
   Debug.Print
   
   ' PCRE Test
   Debug.Print "PCRE Test"
      
   Set lo_RegEx2 = New cPcre2
   With lo_RegEx2.Options.Compile
      .CaseSensitive = False
      .MultiLine = True
   End With
      
   With lo_RegEx2.Options.General
      .GlobalSearch = True
   End With

   Set lo_Matches2 = lo_RegEx2.Execute(l_SubjectText, l_Regex)
   
   Debug.Print "Match Count: " & lo_Matches2.Count
   
   For ii = 0 To lo_Matches2.Count - 1
      Set lo_Match2 = lo_Matches2(ii)
      
      Debug.Print "Match #" & ii + 1 & ": " & lo_Match2.MatchedText
      Debug.Print "Sub Match Count: " & lo_Match2.SubMatchCount
      For jj = 0 To lo_Match2.SubMatchCount - 1
         Debug.Print "SubMatch # " & jj + 1 & ": " & lo_Match2.SubMatchValue(jj)
      Next jj
   Next ii
   Debug.Print
End Sub

Public Function testRunMatch(po_RegexObject As Object, ByVal p_Subject As String, ByVal p_Regex As String, ByVal p_Global As Boolean, ByVal p_IgnoreCase As Boolean) As String
   ' Return log of results
   
   Dim lo_Matches As Object
   Dim lo_Match As Object
   Dim l_Match As String
   Dim l_SubMatch As Long
   Dim l_SubMatchCount As Long
   Dim l_Log As String
   Dim ii As Long
   
   l_Log = vbNewLine & "---------------------------------------------" & vbNewLine
   l_Log = l_Log & "Running testRunMatch test." & vbNewLine
   l_Log = l_Log & "---------------------------------------------" & vbNewLine & vbNewLine
   
   l_Log = l_Log & "Subject: " & p_Subject & vbNewLine
   l_Log = l_Log & "Regex: " & p_Regex & vbNewLine
   l_Log = l_Log & "Is Global: " & p_Global & vbNewLine
   l_Log = l_Log & "Ignore Case: " & p_IgnoreCase & vbNewLine & vbNewLine
   
   With po_RegexObject
      .Pattern = p_Regex
      
      If TypeOf po_RegexObject Is cPcre2 Then
         .GlobalSearch = p_Global
      Else
         .Global = p_Global
      End If
      
      .IgnoreCase = p_IgnoreCase
      
      Set lo_Matches = .Execute(p_Subject)
   End With

   l_Log = l_Log & "Matches Count: " & lo_Matches.Count & vbNewLine & vbNewLine
   
   For Each lo_Match In lo_Matches
      If TypeOf lo_Match Is cPcre2Match Then
         l_Match = lo_Match.MatchedText
      Else
         l_Match = lo_Match.Value
      End If
      
      l_Log = l_Log & "Matched Text: " & l_Match & vbNewLine & vbNewLine
            
      l_Log = l_Log & "Match Start Index: " & lo_Match.FirstIndex & vbNewLine & vbNewLine
      
      If TypeOf lo_Match Is cPcre2Match Then
         l_SubMatchCount = lo_Match.SubMatchCount
      Else
         l_SubMatchCount = lo_Match.SubMatches.Count
      End If
      
      l_Log = l_Log & "Sub-match Count: " & l_SubMatchCount & vbNewLine
      
      For ii = 0 To l_SubMatchCount - 1
         If TypeOf lo_Match Is cPcre2Match Then
            l_SubMatch = lo_Match.SubMatchValue(ii)
         Else
            l_SubMatch = lo_Match.SubMatches(ii)
         End If
         
         l_Log = l_Log & "Sub-match #" & l_SubMatchCount + 1 & ": " & l_SubMatch & vbNewLine
      Next ii
      
      l_Log = l_Log & vbNewLine
   Next lo_Match
   
   l_Log = l_Log & "---------------------------------------------" & vbNewLine & vbNewLine
   
   testRunMatch = l_Log
End Function

Public Function testRunReplace(po_RegexObject As Object, ByVal p_Subject As String, ByVal p_Regex As String, ByVal p_ReplaceWithText As String, ByVal p_Global As Boolean, ByVal p_IgnoreCase As Boolean) As String
   ' Return log of results
   
   Dim l_Result As String
   Dim l_Log As String
   
   l_Log = vbNewLine & "---------------------------------------------" & vbNewLine
   l_Log = l_Log & "Running testRunReplace test." & vbNewLine
   l_Log = l_Log & "---------------------------------------------" & vbNewLine & vbNewLine
   
   l_Log = l_Log & "Subject: " & p_Subject & vbNewLine
   l_Log = l_Log & "Regex: " & p_Regex & vbNewLine
   l_Log = l_Log & "Replace with Text: " & p_ReplaceWithText & vbNewLine
   l_Log = l_Log & "Is Global: " & p_Global & vbNewLine
   l_Log = l_Log & "Ignore Case: " & p_IgnoreCase & vbNewLine & vbNewLine
   
   With po_RegexObject
      .Pattern = p_Regex
      
      If TypeOf po_RegexObject Is cPcre2 Then
         .GlobalSearch = p_Global
      Else
         .Global = p_Global
      End If
      
      .IgnoreCase = p_IgnoreCase
      
      l_Result = .Replace(p_Subject, p_ReplaceWithText)
   End With

   l_Log = l_Log & "Replacement Result: " & l_Result & vbNewLine & vbNewLine
   
   l_Log = l_Log & "---------------------------------------------" & vbNewLine & vbNewLine
   
   testRunReplace = l_Log
End Function


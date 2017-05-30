Attribute VB_Name = "modTests"
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

Sub TestRegexReplace()
   Dim lo_RegEx As New CPcre
   Dim lo_Matches As CPcreMatches
   Dim ii As Long

   With lo_RegEx.Options.Compile
      .CaseSensitive = False
   End With
   With lo_RegEx.Options.General
      .GlobalSearch = True
   End With
   
   Debug.Print "Replace result: " & lo_RegEx.Replace("This is a test of matching test stuff!", "XXXXXXX", "test")
End Sub

Sub TestRegexMatch()
   Dim lo_RegEx As New CPcre
   Dim lo_Matches As CPcreMatches
   Dim lo_Match As CPcreMatch
   Dim ii As Long

   With lo_RegEx.Options.Compile
      .CaseSensitive = False
   End With
   
   Set lo_Matches = lo_RegEx.Execute("This is a test of matching stuff!", "(test)\s*.+\s*(Mat)")
   If lo_Matches.Count > 0 Then
      For ii = 0 To lo_Matches.Count - 1
         Debug.Print "Match #" & ii + 1 & ": " & lo_Matches(ii).MatchedText
      Next ii
      
   Else
      Debug.Print "No matches found!"
   End If
End Sub

Sub TestRegexCallout(po_Pcre As CPcre)
   Dim lo_Matches As CPcreMatches
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

Sub TestRegexMatchedEvent(po_Pcre As CPcre)
   Dim lo_Matches As CPcreMatches
   
   With po_Pcre.Options.Compile
      .CaseSensitive = False
   End With
   
   With po_Pcre.Options.General
      .GlobalSearch = True
   End With
   
   With po_Pcre.Options.Match
      .MatchedEventEnabled = True
   End With
   
   Set lo_Matches = po_Pcre.Execute("This is a test of CB9CBC213211BF4BA026FED4B1AC5CB2. Did you know that CB9CBC213211BF4BA026FED4B1AC5CB2 can be cached for performance, or re-run for each match. F8CEC6F354497746990BFB3A6A72BD06, F8CEC6F354497746990BFB3A6A72BD06, F8CEC6F354497746990BFB3A6A72BD06. You can also ignore matches: E175FA8438E00D47B2AA52CAD413FB6A.", "[a-fA-F0-9]{32}")

   Debug.Print "Result text: " & lo_Matches.Text
End Sub

Sub TestRegex2()
   Dim lo_RegEx As Object ' VBScript_RegExp_55.RegExp
   Dim lo_Matches As Object 'VBScript_RegExp_55.MatchCollection
   Dim lo_Match As Object 'VBScript_RegExp_55.Match
   
   Dim lo_RegEx2 As CPcre
   Dim lo_Matches2 As CPcreMatches
   Dim lo_Match2 As CPcreMatch

   Dim l_SubjectText As String
   Dim l_Regex As String
   
   Dim ii As Long
   Dim jj As Long
   
   l_SubjectText = "File1.zip.exe" & vbCrLf & "File2.com" & vbCrLf & "File 3"
   l_Regex = "[\w ]+(\.\S+?)*$"
   
   ' VBScript Test
   Debug.Print "VBSCRIPT Test"
   
   Set lo_RegEx = CreateObject("VBScript.RegExp")
   With lo_RegEx
      .IgnoreCase = True
      .Global = True
      .Multiline = True
   End With
   
   lo_RegEx.Pattern = l_Regex
   
   Set lo_Matches = lo_RegEx.Execute(l_SubjectText)
   
   Debug.Print "Match Count: " & lo_Matches.Count
         
   For Each lo_Match In lo_Matches
      Debug.Print "Match #" & ii + 1 & ": " & lo_Match.Value
      Debug.Print "Sub Match Count: " & lo_Match.SubMatches.Count
      For jj = 0 To lo_Match.SubMatches.Count - 1
         Debug.Print "SubMatch # " & jj + 1 & ": " & lo_Match.SubMatches(jj)
      Next jj
   Next lo_Match
   Debug.Print
   
   ' PCRE Test
   Debug.Print "PCRE Test"
      
   Set lo_RegEx2 = New CPcre
   With lo_RegEx2.Options.Compile
      .CaseSensitive = False
      .Multiline = True
   End With
      
   With lo_RegEx2.Options.General
      .GlobalSearch = True
   End With

   Set lo_Matches2 = lo_RegEx2.Execute(l_SubjectText, l_Regex)
   
   Debug.Print "Match Count: " & lo_Matches2.Count
   
   For Each lo_Match2 In lo_Matches2
      Debug.Print "Match #" & ii + 1 & ": " & lo_Match2.MatchedText
      Debug.Print "Sub Match Count: " & lo_Match2.SubMatchCount
      For jj = 0 To lo_Match2.SubMatchCount - 1
         Debug.Print "SubMatch # " & jj + 1 & ": " & lo_Match2.SubMatchValue(jj)
      Next jj
   Next lo_Match2
   Debug.Print
End Sub

Public Function testRunMatch(po_RegexObject As Object, ByVal p_Subject As String, ByVal p_Regex As String, ByVal p_Global As Boolean, ByVal p_IgnoreCase As Boolean) As String
   ' Return log of results
   
   Dim lo_Matches As Object
   Dim lo_Match As Object
   Dim l_Match As String
   Dim ii As Long
   Dim l_Log As String
   
   l_Log = vbNewLine & "---------------------------------------------" & vbNewLine
   l_Log = l_Log & "Running testRunMatch test." & vbNewLine
   l_Log = l_Log & "---------------------------------------------" & vbNewLine & vbNewLine
   
   l_Log = l_Log & "Subject: " & p_Subject & vbNewLine
   l_Log = l_Log & "Regex: " & p_Regex & vbNewLine
   l_Log = l_Log & "Is Global: " & p_Global & vbNewLine
   l_Log = l_Log & "Ignore Case: " & p_IgnoreCase & vbNewLine & vbNewLine
   
   With po_RegexObject
      .Pattern = p_Regex
      
      If TypeOf po_RegexObject Is CPcre Then
         .GlobalSearch = p_Global
      Else
         .Global = p_Global
      End If
      
      .IgnoreCase = p_IgnoreCase
      
      Set lo_Matches = .Execute(p_Subject)
   End With

   l_Log = l_Log & "Matches Count: " & lo_Matches.Count & vbNewLine & vbNewLine
   
   For Each lo_Match In lo_Matches
      If TypeOf lo_Match Is CPcreMatch Then
         l_Match = lo_Match.MatchedText
      Else
         l_Match = lo_Match.Value
      End If
      
      l_Log = l_Log & "Matched Text: " & l_Match & vbNewLine & vbNewLine
   Next lo_Match
   
   l_Log = l_Log & "---------------------------------------------" & vbNewLine & vbNewLine
   
   testRunMatch = l_Log
End Function

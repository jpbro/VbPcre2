Attribute VB_Name = "modTests"
Option Explicit

Sub TestRegex()
   Dim lo_RegEx As New CPcre
   Dim lo_Matches As CPcreMatches
   Dim ii As Long
   
   With lo_RegEx.RegexOptions
      .CaseSensitive = False
   End With
   
   Set lo_Matches = lo_RegEx.Execute("This is a test of matching stuff!", "(test)\s*.+\s*(Mat)")
   If lo_Matches.Count > 0 Then
      For ii = 0 To lo_Matches.Count - 1
         Debug.Print "Match #" & ii + 1 & ": " & lo_Matches(ii).Value
      Next ii
      
   Else
      Debug.Print "No matches found!"
   End If
End Sub

Sub TestRegexCallout()
   Dim lo_RegEx As New CPcre
   Dim lo_Matches As CPcreMatches
   Dim ii As Long
   
   With lo_RegEx.RegexOptions
      .CaseSensitive = False
   End With
   
   Set lo_Matches = lo_RegEx.Execute("This is a test of matching stuff!", "(?C1)(test)\s*.+\s*(Mat)")
   If lo_Matches.Count > 0 Then
      For ii = 0 To lo_Matches.Count - 1
         Debug.Print "Match #" & ii + 1 & ": " & lo_Matches(ii).Value
      Next ii
      
   Else
      Debug.Print "No matches found!"
   End If
End Sub

Sub TestRegex2()
   Dim lo_RegEx As VBScript_RegExp_55.RegExp
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
   With lo_RegEx2.RegexOptions
      .CaseSensitive = False
      .GlobalSearch = True
      .Multiline = True
   End With

   Set lo_Matches2 = lo_RegEx2.Execute(l_SubjectText, l_Regex)
   
   Debug.Print "Match Count: " & lo_Matches2.Count
   
   For Each lo_Match2 In lo_Matches2
      Debug.Print "Match #" & ii + 1 & ": " & lo_Match2.Value
      Debug.Print "Sub Match Count: " & lo_Match2.SubMatchCount
      For jj = 0 To lo_Match2.SubMatchCount - 1
         Debug.Print "SubMatch # " & jj + 1 & ": " & lo_Match2.SubMatchValue(jj)
      Next jj
   Next lo_Match2
   Debug.Print
End Sub

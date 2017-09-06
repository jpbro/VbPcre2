VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "VbPcre2 Test"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   11600
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtbTestResults 
      Height          =   3252
      Index           =   0
      Left            =   144
      TabIndex        =   1
      Top             =   1440
      Width           =   5052
      _ExtentX        =   8908
      _ExtentY        =   5733
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton cmdRunTests 
      Caption         =   "Run Tests"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   4392
      TabIndex        =   0
      Top             =   216
      Width           =   2748
   End
   Begin RichTextLib.RichTextBox rtbTestResults 
      Height          =   3252
      Index           =   1
      Left            =   5904
      TabIndex        =   2
      Top             =   1476
      Width           =   5052
      _ExtentX        =   8908
      _ExtentY        =   5733
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form1.frx":007B
   End
   Begin VB.Label Label 
      Caption         =   "VB Script"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   5976
      TabIndex        =   6
      Top             =   1080
      Width           =   1560
   End
   Begin VB.Label Label 
      Caption         =   "VbPcre2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   216
      TabIndex        =   5
      Top             =   1080
      Width           =   1560
   End
   Begin VB.Label lblRunTime 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Index           =   1
      Left            =   6012
      TabIndex        =   4
      Top             =   4932
      Width           =   4800
   End
   Begin VB.Label lblRunTime 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   4896
      Width           =   4800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private WithEvents mo_Pcre As cPcre2
Attribute mo_Pcre.VB_VarHelpID = -1

Private Sub cmdRunTests_Click()
   Dim l_Subject As String
   Dim l_Regex As String
   Dim l_Replace As String
   
   ' Matches should be found in these test. Results depend on Global and IgnoreCase parameter values
   l_Subject = "This is a test of the emergency broadcast system." & vbNewLine & "This is only a TEST!" & vbNewLine & "If this were not a test, then really bad things have happened."
   l_Regex = "test"
   
   Me.rtbTestResults(0).SelStart = 0
   Me.rtbTestResults(1).SelStart = 0
   
   Me.rtbTestResults(0).Text = testRunMatch(New cPcre2, l_Subject, l_Regex, False, False)
   Me.rtbTestResults(1).Text = testRunMatch(CreateObject("VBScript.Regexp"), l_Subject, l_Regex, False, False)
   
   Me.rtbTestResults(0).Text = Me.rtbTestResults(0).Text & testRunMatch(New cPcre2, l_Subject, l_Regex, True, False)
   Me.rtbTestResults(1).Text = Me.rtbTestResults(1).Text & testRunMatch(CreateObject("VBScript.Regexp"), l_Subject, l_Regex, True, False)

   Me.rtbTestResults(0).Text = Me.rtbTestResults(0).Text & testRunMatch(New cPcre2, l_Subject, l_Regex, True, True)
   Me.rtbTestResults(1).Text = Me.rtbTestResults(1).Text & testRunMatch(CreateObject("VBScript.Regexp"), l_Subject, l_Regex, True, True)

   Me.rtbTestResults(0).Text = Me.rtbTestResults(0).Text & testRunMatch(New cPcre2, l_Subject, l_Regex, False, True)
   Me.rtbTestResults(1).Text = Me.rtbTestResults(1).Text & testRunMatch(CreateObject("VBScript.Regexp"), l_Subject, l_Regex, False, True)

   ' No matches should be found in these following tests
   l_Regex = "gobbledeegook"

   Me.rtbTestResults(0).Text = Me.rtbTestResults(0).Text & testRunMatch(New cPcre2, l_Subject, l_Regex, False, False)
   Me.rtbTestResults(1).Text = Me.rtbTestResults(1).Text & testRunMatch(CreateObject("VBScript.Regexp"), l_Subject, l_Regex, False, False)

   Me.rtbTestResults(0).Text = Me.rtbTestResults(0).Text & testRunMatch(New cPcre2, l_Subject, l_Regex, True, False)
   Me.rtbTestResults(1).Text = Me.rtbTestResults(1).Text & testRunMatch(CreateObject("VBScript.Regexp"), l_Subject, l_Regex, True, False)

   Me.rtbTestResults(0).Text = Me.rtbTestResults(0).Text & testRunMatch(New cPcre2, l_Subject, l_Regex, True, True)
   Me.rtbTestResults(1).Text = Me.rtbTestResults(1).Text & testRunMatch(CreateObject("VBScript.Regexp"), l_Subject, l_Regex, True, True)

   Me.rtbTestResults(0).Text = Me.rtbTestResults(0).Text & testRunMatch(New cPcre2, l_Subject, l_Regex, False, True)
   Me.rtbTestResults(1).Text = Me.rtbTestResults(1).Text & testRunMatch(CreateObject("VBScript.Regexp"), l_Subject, l_Regex, False, True)

   ' Substitution Tests
   l_Regex = "test"
   l_Replace = "<REDACTED>"

   Me.rtbTestResults(0).Text = Me.rtbTestResults(0).Text & testRunReplace(New cPcre2, l_Subject, l_Regex, l_Replace, False, False)
   Me.rtbTestResults(1).Text = Me.rtbTestResults(1).Text & testRunReplace(CreateObject("VBScript.Regexp"), l_Subject, l_Regex, l_Replace, False, False)

   Me.rtbTestResults(0).Text = Me.rtbTestResults(0).Text & testRunReplace(New cPcre2, l_Subject, l_Regex, l_Replace, True, False)
   Me.rtbTestResults(1).Text = Me.rtbTestResults(1).Text & testRunReplace(CreateObject("VBScript.Regexp"), l_Subject, l_Regex, l_Replace, True, False)

   Me.rtbTestResults(0).Text = Me.rtbTestResults(0).Text & testRunReplace(New cPcre2, l_Subject, l_Regex, l_Replace, True, True)
   Me.rtbTestResults(1).Text = Me.rtbTestResults(1).Text & testRunReplace(CreateObject("VBScript.Regexp"), l_Subject, l_Regex, l_Replace, True, True)

   Me.rtbTestResults(0).Text = Me.rtbTestResults(0).Text & testRunReplace(New cPcre2, l_Subject, l_Regex, l_Replace, False, True)
   Me.rtbTestResults(1).Text = Me.rtbTestResults(1).Text & testRunReplace(CreateObject("VBScript.Regexp"), l_Subject, l_Regex, l_Replace, False, True)

   Set mo_Pcre = New cPcre2
   TestRegexEnumerateCallout mo_Pcre
   TestRegexCallout mo_Pcre
      
   ' Report Results
   If Me.rtbTestResults(0).Text = Me.rtbTestResults(1).Text Then
      MsgBox "Test results match :)", vbOKOnly, "Result Match"
   Else
      MsgBox "Test results DO NOT MATCH!", vbExclamation + vbOKOnly, "Result Mismatch!"
   End If
End Sub

Private Sub Form_Resize()
   Me.cmdRunTests.Left = Me.ScaleWidth / 2 - Me.cmdRunTests.Width / 2
   
   With Me.rtbTestResults(0)
      .Move .Left, .Top, Me.ScaleWidth / 2 - .Left * 2, Me.ScaleHeight - .Top - Me.lblRunTime(0).Height - .Left * 3
   End With

   With Me.rtbTestResults(1)
      .Move Me.ScaleWidth - Me.rtbTestResults(0).Width - Me.rtbTestResults(0).Left, Me.rtbTestResults(0).Top, Me.rtbTestResults(0).Width, Me.rtbTestResults(0).Height
   End With

   With Me.lblRunTime(0)
      .Move Me.rtbTestResults(0).Left, Me.ScaleHeight - .Height - Me.rtbTestResults(0).Left, Me.rtbTestResults(0).Width, .Height
   End With

   With Me.lblRunTime(1)
      .Move Me.rtbTestResults(1).Left, Me.ScaleHeight - .Height - Me.rtbTestResults(0).Left, Me.rtbTestResults(1).Width, .Height
   End With
End Sub

Private Sub mo_Pcre_CalloutEnumerated(ByVal p_CalloutNumber As Long, ByVal p_CalloutLabel As String, ByVal p_CalloutOffset As Long, ByVal p_PatternPosition As Long, ByVal p_NextItemLength As Long, p_Action As e_CalloutEnumeratedAction)
   Debug.Print "Callout #" & p_CalloutNumber & "'" & p_CalloutLabel & "' Enumerated in " & Me.Name

End Sub

Private Sub mo_Pcre_CalloutReceived(ByVal p_CalloutNumber As Long, ByVal p_CalloutLabel As String, ByVal p_CalloutOffset As Long, ByVal p_Subject As String, ByVal p_Mark As String, ByVal p_CaptureTop As Long, ByVal p_CaptureLast As Long, pa_OffsetVector() As Long, ByVal p_PatternPosition As Long, ByVal p_NextItemLength As Long, p_Action As e_CalloutReceivedAction)
   Debug.Print "Callout #" & p_CalloutNumber & " named '" & p_CalloutLabel & "' Received in " & Me.Name
End Sub

Private Sub mo_Pcre_Matched(p_MatchedText As String, p_SubstitutionAction As e_SubstitutionAction, p_Cancel As Boolean)
   Static s_MatchCount As Long
   
   Debug.Print "MATCH FOUND: " & p_MatchedText
   
   Select Case p_MatchedText
   Case "CB9CBC213211BF4BA026FED4B1AC5CB2"
      p_MatchedText = "match substitution"
      p_SubstitutionAction = subaction_ReplaceAndCache
      
   Case "F8CEC6F354497746990BFB3A6A72BD06"
      s_MatchCount = s_MatchCount + 1
      p_MatchedText = s_MatchCount
      p_SubstitutionAction = subaction_Replace
   End Select
End Sub

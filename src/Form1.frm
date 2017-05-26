VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3024
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3024
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Private WithEvents mo_Pcre As CPcre
Attribute mo_Pcre.VB_VarHelpID = -1

Private Sub Form_Load()
   Set mo_Pcre = New CPcre
   
   TestRegexCallout mo_Pcre
   Unload Me
End Sub

Private Sub mo_Pcre_CalloutEnumerateReceived(ByVal p_CalloutNumber As Long, ByVal p_PatternPosition As Long, ByVal p_NextItemLength As Long, ByVal p_CalloutOffset As Long, ByVal p_CalloutLength As Long, ByVal p_CalloutString As String, p_ReturnValue As Long)
   Debug.Print "Callout #" & p_CalloutNumber & " Received in " & Me.Name
End Sub

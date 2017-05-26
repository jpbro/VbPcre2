VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPcreMatches"
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

Private mo_Matches As VBA.Collection

Public Property Get Count() As Long
   Count = mo_Matches.Count
End Property

Public Property Get Item(ByVal p_ZeroBasedIndex As Long) As CPcreMatch
Attribute Item.VB_UserMemId = 0
   Set Item = mo_Matches.Item(p_ZeroBasedIndex + 1)
End Property

Public Function Enumerator() As IUnknown
Attribute Enumerator.VB_UserMemId = -4
Attribute Enumerator.VB_MemberFlags = "40"
   Set Enumerator = mo_Matches.[_NewEnum]
End Function

Friend Sub Add(ByVal p_TextToSearch As String, ByVal p_OvectorPointer As Long, ByVal p_MatchCount As Long)
   Dim lo_Match As CPcreMatch
   
   Set lo_Match = New CPcreMatch
   
   lo_Match.Init p_TextToSearch, p_OvectorPointer, p_MatchCount
   
   mo_Matches.Add lo_Match
End Sub

Private Sub Class_Initialize()
   Set mo_Matches = New VBA.Collection
End Sub

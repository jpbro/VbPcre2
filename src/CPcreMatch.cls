VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPcreMatch"
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

Private m_TextToSearch As String
Private ma_Ovector() As Long
Private m_Count As Long

Public Property Get Value() As String
   Value = Me.SubMatchValue(-1)
End Property

Public Property Get SubMatchCount() As Long
   If m_Count > 1 Then
      SubMatchCount = m_Count - 1
   End If
End Property

Public Property Get SubMatchValue(ByVal p_ZeroBasedSubMatchIndex As Long) As String
   p_ZeroBasedSubMatchIndex = 2 * (p_ZeroBasedSubMatchIndex + 1)
   
   SubMatchValue = Mid$(m_TextToSearch, ma_Ovector(p_ZeroBasedSubMatchIndex) + 1, ma_Ovector(p_ZeroBasedSubMatchIndex + 1) - ma_Ovector(p_ZeroBasedSubMatchIndex))
End Property

Friend Sub Init(ByVal p_TextToSearch As String, ByVal p_OvectorPointer As Long, ByVal p_MatchCount As Long)
   If p_MatchCount > 0 Then
      ReDim ma_Ovector(2 * p_MatchCount - 1)
   
      CopyMemory ma_Ovector(0), ByVal p_OvectorPointer, 2 * p_MatchCount * 4
   End If
   
   m_TextToSearch = p_TextToSearch
   
   m_Count = p_MatchCount
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPcre2Match"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private m_SearchedText As String
Private ma_Ovector() As Long
Private m_Count As Long
Private m_MatchOffset As Long

Friend Property Get SearchedText() As String
   ' The full text searched up to the end of the matched text
   SearchedText = m_SearchedText
End Property

Public Property Get MatchedText() As String
   ' The matched chunk of text in the searched text
   MatchedText = Me.SubMatchValue(-1)
End Property

Public Property Get FirstIndex() As Long
   ' Return the character index of the start of the full match
   FirstIndex = ma_Ovector(0) + m_MatchOffset
End Property

Public Property Get LastIndex() As Long
   ' Return the character index of the end of the full match
   LastIndex = ma_Ovector(1) + m_MatchOffset
End Property

Public Property Get Length() As Long
   ' Return the length of the full match
   Length = ma_Ovector(1) - ma_Ovector(0)
End Property

Public Property Get SubMatchCount() As Long
   If m_Count > 1 Then
      SubMatchCount = m_Count - 1
   End If
End Property

Public Property Get SubMatchValue(ByVal p_ZeroBasedSubMatchIndex As Long) As String
   Dim l_Len As Long
   
   p_ZeroBasedSubMatchIndex = 2 * (p_ZeroBasedSubMatchIndex + 1)
   
   l_Len = ma_Ovector(p_ZeroBasedSubMatchIndex + 1) - ma_Ovector(p_ZeroBasedSubMatchIndex)

   If l_Len > 0 Then
      SubMatchValue = Mid$(m_SearchedText, ma_Ovector(p_ZeroBasedSubMatchIndex) + 1, l_Len)
   End If
End Property

Public Property Get SubMatchFirstIndex(ByVal p_ZeroBasedSubMatchIndex As Long) As Long
   p_ZeroBasedSubMatchIndex = 2 * (p_ZeroBasedSubMatchIndex + 1)
   
   SubMatchFirstIndex = ma_Ovector(p_ZeroBasedSubMatchIndex)
End Property

Public Property Get SubMatchLastIndex(ByVal p_ZeroBasedSubMatchIndex As Long) As Long
   p_ZeroBasedSubMatchIndex = 2 * (p_ZeroBasedSubMatchIndex + 1)
   
   SubMatchLastIndex = ma_Ovector(p_ZeroBasedSubMatchIndex + 1)
End Property

Friend Sub Init(ByVal p_SearchedText As String, pa_OvectorArray() As Long, ByVal p_MatchCount As Long, ByVal p_MatchOffset As Long)
   If p_MatchCount > 0 Then
      ma_Ovector = pa_OvectorArray
   End If
   
   m_SearchedText = p_SearchedText
   
   m_Count = p_MatchCount
   
   m_MatchOffset = p_MatchOffset
End Sub


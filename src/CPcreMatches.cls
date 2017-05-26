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

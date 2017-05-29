VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPcreOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Replace() As CPcreOptionsReplace
   Static so_Options As CPcreOptionsReplace

   If so_Options Is Nothing Then
      Set so_Options = New CPcreOptionsReplace
   End If
   Set Replace = so_Options
End Function

Public Function Match() As CPcreOptionsMatch
   Static so_Options As CPcreOptionsMatch

   If so_Options Is Nothing Then
      Set so_Options = New CPcreOptionsMatch
   End If
   Set Match = so_Options
End Function


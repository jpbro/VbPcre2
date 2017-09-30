VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPcre2Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Compile() As cPcre2OptionsCompile
   Static so_Options As cPcre2OptionsCompile

   If so_Options Is Nothing Then
      Set so_Options = New cPcre2OptionsCompile
   End If
   Set Compile = so_Options
End Function

Public Function General() As cPcre2OptionsGeneral
   Static so_Options As cPcre2OptionsGeneral

   If so_Options Is Nothing Then
      Set so_Options = New cPcre2OptionsGeneral
   End If
   Set General = so_Options
End Function

Public Function Match() As cPcre2OptionsMatch
   Static so_Options As cPcre2OptionsMatch

   If so_Options Is Nothing Then
      Set so_Options = New cPcre2OptionsMatch
   End If
   Set Match = so_Options
End Function

Public Function Replace() As cPcre2OptionsReplace
   Static so_Options As cPcre2OptionsReplace

   If so_Options Is Nothing Then
      Set so_Options = New cPcre2OptionsReplace
   End If
   Set Replace = so_Options
End Function


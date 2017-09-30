Attribute VB_Name = "mWin32"
Option Explicit

Public Declare Sub win32_CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Function win32_LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal p_LibraryFileName As String) As Long
Public Declare Function win32_FreeLibrary Lib "kernel32.dll" Alias "FreeLibrary" (ByVal p_Hmodule As Long) As Long


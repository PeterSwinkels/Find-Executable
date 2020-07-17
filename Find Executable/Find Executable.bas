Attribute VB_Name = "FindExecutableModule"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API constants and functions used by this program.
Private Const SE_ERR_ACCESSDENIED As Long = 5   'The specified file cannot be accessed.
Private Const SE_ERR_FNF As Long = 2            'The specified file was not found.
Private Const SE_ERR_NOASSOC As Long = 31       'There is no association for the specified file type with an executable file.
Private Const SE_ERR_OOM As Long = 8            'The system is out of memory or resources.
Private Const SE_ERR_PNF As Long = 3            'The specified path is invalid.

Private Declare Function FindExecutableA Lib "Shell32.dll" (ByVal lpFile As String, ByVallpDirectory As String, ByVal lpResult As String) As Long

'The constants used by this program.
Private Const MAX_PATH As Long = 256   'Defines the maximum length allowed for paths.
'This procedure retrieves and returns the executable associated with the specified file.
Private Function FindExecutableName(FileName As String) As String
Dim ExecutableName As String
Dim ReturnValue As Long

   ExecutableName = String$(MAX_PATH, vbNullChar)
   ReturnValue = FindExecutableA(FileName, vbNullString, ExecutableName)
   If ReturnValue <= 32 Then ExecutableName = vbNullString
   If InStr(ExecutableName, vbNullChar) > 0 Then ExecutableName = Left$(ExecutableName, InStr(ExecutableName, vbNullChar) - 1)

   FindExecutableName = ExecutableName
End Function

'This procedure is executed when this program is started.
Public Sub Main()
Dim ExecutableName As String
Dim FileName As String

   FileName = InputBox$("Enter a file name:")
   If FileName = vbNullString Then Exit Sub

   ExecutableName = FindExecutableName(FileName)

   If ExecutableName = vbNullString Then
      MsgBox "Could not find a file association.", vbExclamation
   Else
      MsgBox "This file is opened by: " & vbCr & """" & ExecutableName & """", vbInformation
   End If
End Sub



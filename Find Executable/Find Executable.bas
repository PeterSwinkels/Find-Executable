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
'This procedure retrieves and returns the path associated with the specified executable.
Private Function FindExecutablePath(FileName As String) As String
On Error GoTo ErrorTrap
Dim ExecutablePath As String
Dim ReturnValue As Long

   ExecutablePath = String$(MAX_PATH, vbNullChar)
   ReturnValue = FindExecutableA(FileName, vbNullString, ExecutablePath)
   If ReturnValue <= SE_ERR_NOASSOC Then ExecutablePath = vbNullString
   If InStr(ExecutablePath, vbNullChar) > 0 Then ExecutablePath = Left$(ExecutablePath, InStr(ExecutablePath, vbNullChar) - 1)

EndProcedure:
   FindExecutablePath = ExecutablePath
   Exit Function
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbRetry Then Resume
End Function

'This procedure handles any errors that occur.
Private Function HandleError(Optional ReturnPreviousChoice As Boolean = False) As Long
Dim Description As String
Dim ErrorCode As Long
Static Choice As Long

   Description = Err.Description
   ErrorCode = Err.Number
   On Error Resume Next
   If Not ReturnPreviousChoice Then
      Choice = MsgBox(Description & "." & vbCr & "Error code: " & CStr(ErrorCode), vbAbortRetryIgnore Or vbDefaultButton2 Or vbExclamation)
   End If
   
   If Choice = vbAbort Then End
   
   HandleError = Choice
End Function
'This procedure is executed when this program is started.
Public Sub Main()
On Error GoTo ErrorTrap
Dim ExecutableName As String
Dim FileName As String

   FileName = InputBox$("Enter a file name:")
   If Not FileName = vbNullString Then
      ExecutableName = FindExecutablePath(FileName)
   
      If ExecutableName = vbNullString Then
         MsgBox "Could not find a file association.", vbExclamation
      Else
         MsgBox "This file is opened by: " & vbCr & """" & ExecutableName & """", vbInformation
      End If
   End If

EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbRetry Then Resume
End Sub



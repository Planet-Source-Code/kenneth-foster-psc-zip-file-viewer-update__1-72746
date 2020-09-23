Attribute VB_Name = "modKillFolder"
'Code By Rde
Option Explicit

Private Declare Function GetAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpSpec As String) As Long
Private Declare Function SetAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpSpec As String, ByVal dwAttributes As Long) As Long

Private Const DIR_SEP As String = "\"
Private Const INVALID_FILE_ATTRIBUTES = (-1)

'-----------------------------------------------------

' This is a Kill Folder function with persistence.

' It will remove all sub-folders and files and then
' optionally delete the specified folder.

' I found when removing all the files in the temp folder
' that some locked files would fail and cause it to not
' continue with the rest of the files.

' This function will continue to remove all unlocked files,
' even after finding locked files. However, if locked files
' are found, the parent folder will also not get removed.

'-----------------------------------------------------

Public Function AddBackslash(sPath As String) As String
If Right$(sPath, 1&) = DIR_SEP Then
AddBackslash = sPath
Else
AddBackslash = sPath & DIR_SEP
End If
End Function


'-----------------------------------------------------

Public Function FolderExists(sPath As String) As Boolean

Dim Attribs As Long
Attribs = GetAttributes(sPath)
If Not (Attribs = INVALID_FILE_ATTRIBUTES) Then
FolderExists = ((Attribs And vbDirectory) = vbDirectory)
End If
End Function


'-----------------------------------------------------

Public Function KillFolder(sSpec As String, Optional ByVal bJustEmptyDontRemove As Boolean) As Boolean

Dim sRoot As String, sDir As String, sFile As String
Dim iCnt As Long, iIdx As Long

If Not FolderExists(sSpec) Then Exit Function

' Add trailing backslash if missing
sRoot = AddBackslash(sSpec)
iCnt = 2& '.' '..'

On Error Resume Next ' Ignore file errors
sFile = Dir$(sRoot & "*.*", vbNormal)
Do While LenB(sFile)
SetAttributes sRoot & sFile, vbNormal
Kill sRoot & sFile
sFile = Dir$
Loop

On Error GoTo HandleIt ' No error should occur in here
Do: sDir = Dir$(sRoot & "*", vbDirectory)
For iIdx = 1& To iCnt
sDir = Dir$ '.' '..' ['fail']
Next
If LenB(sDir) = 0& Then Exit Do
If KillFolder(sRoot & sDir & DIR_SEP) Then
' Sub-folder is now gone but Dir$ was reset
' during recursive call so Do Dir$(..) again
Else: iCnt = iCnt + 1&
' Kill folder failed (remnant files) so skip
' this folder (iCnt + 1) to get the rest
End If
Loop

If bJustEmptyDontRemove = False Then RmDir sRoot ' Errors here if remnants
HandleIt:
KillFolder = Not FolderExists(sSpec)
End Function


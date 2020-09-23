Attribute VB_Name = "modMCI"
Option Explicit

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Const SND_ASYNC As Long = &H1
Public Const SND_NODEFAULT As Long = &H2
Public Const SND_MEMORY As Long = &H4
Public Const SND_LOOP As Long = &H8
Public Const SND_NOSTOP As Long = &H10
    
Public Sub PlaySound(ByVal strFileName As String)
    DoEvents
    Call sndPlaySound(strFileName, SND_ASYNC)
    DoEvents
End Sub


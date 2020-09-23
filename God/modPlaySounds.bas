Attribute VB_Name = "modPlaySounds"
Private Declare Function PlaySound& Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long)

Private Const SND_ALIAS& = &H10000
' Playsound returns immediately
' Do not use SND_SYNC
Private Const SND_ASYNC& = &H1
' The name of a wave file.
' Do not use with SND_RESOURCE or SND_AL
'     IAS
Private Const SND_FILENAME& = &H20000
' Unless used, the default beep will
' play if the specified resource is miss
'     ing
Private Const SND_NODEFAULT& = &H2
' Fail the call & do not wait for
' a sound device if it is otherwise unav
'     ailable
Private Const SND_NOWAIT& = &H2000
' Use a resource file as the source.
' Do not use with SND_ALIAS or SND_FILEN
'     AME
Private Const SND_RESOURCE& = &H40004
' Playsound will not return until the
' specified sound has played. Do not
' use with SND_ASYNC
Private Const SND_SYNC& = &H0

Public Enum enSound_Source
    ssFile = SND_FILENAME&
    ssRegistry = SND_ALIAS&
End Enum

Public Const elDefault = ".Default"
Public Const elGPF = "AppGPFault"
Public Const elClose = "Close"
Public Const elEmptyRecycleBin = "EmptyRecycleBin"
Public Const elMailBeep = "MailBeep"
Public Const elMaximize = "Maximize"
Public Const elMenuCommand = "MenuCommand"
Public Const elMenuPopUp = "MenuPopup"
Public Const elMinimize = "Minimize"
Public Const elOpen = "Open"
Public Const elRestoreDown = "RestoreDown"
Public Const elRestoreUp = "RestoreUp"
Public Const elSystemAsterisk = "SystemAsterisk"
Public Const elSystemExclaimation = "SystemExclaimation"
Public Const elSystemExit = "SystemExit"
Public Const elSystemHand = "SystemHand"
Public Const elSystemQuestion = "SystemQuestion"
Public Const elSystemStart = "SystemStart"

Public Function EZPlay(ssname As String, sound_source As enSound_Source) As Boolean
    
    If PlaySound(ssname, 0&, sound_source Or SND_ASYNC Or SND_NODEFAULT) Then
        EZPlay = True
    Else
        EZPlay = False
    End If

End Function



VERSION 5.00
Begin VB.Form console 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Console"
   ClientHeight    =   3765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   3735
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCurrentTime Lib "kernel32" Alias "GetTickCount" () As Long
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Dim num As Integer
Dim prompt As String
Dim keys(2) As Integer

Private Sub Form_Load()
    Me.WindowState = 2
    num = 1
    Me.Show
    txt.SelStart = Len(txt)
    prompt = "uni.vers.edu"
    
    getNum "s", 0
    getNum "e", 1
    getNum "k", 2
    
    run
End Sub

Private Sub getNum(a As String, b As Integer)
    prefix = a
    i = 0
    tmpStr = Dir(App.Path & "\" & prefix & "*.wav")
    Do While tmpStr <> ""
        tmpStr = Dir
        i = i + 1
    Loop
    keys(b) = i
End Sub

Private Sub Form_Resize()
    txt.Width = Width
    txt.Height = Height
    txt.Left = 0
    txt.Top = 0
End Sub

Private Sub run()
    ms "In the beginning there was the computer. And God said,"
    'pause (2000)
    typ "Let there be light!"
    ms "Not logged in. Log in before entering any commands"
    pB
    typ "user", 2000
    ms "Enter user id."
    typ "God", 500
    ms "Username God accepted. Enter password for God."
    typ "Omniscient"
    ms "Password incorrect. Enter password for God@uni.vers.edu or SysRq to cancel"
    pB
    typ "Omnipotent", 2000
    ms "Password incorrect. Enter password for God@uni.vers.edu or SysRq to cancel"
    pB
    typ "Technocrat", 3000
    ms "And God logged on at 12:01:00 AM, Sunday, March 1."
    prompt = "God@uni.vers.edu"
    typ "Let there be light!", 1500
    ms "Unrecognizable command. Try again."
    typ "Create light"
    ms "Done"
    txt.ForeColor = 0
    txt.BackColor = RGB(255, 255, 255)
    typ "Run heaven_and_earth"
    ms "And God created Day and Night. And God saw there were 0 errors."
    ms "And God logged off at 12:02:00 AM, Sunday, March 1."
    pause (3000)
    txt = "": num = 1
    ms "And God logged on at 12:01:00 AM, Monday, March 2."
    typ "Let there be firmament in the midst of water and light"
    ms "Unrecognizable command. Try again."
    typ "Create firmament", 2000
    ms "Done."
    typ "Run firmament"
    ms "And God divided the waters. And God saw there were 0 errors."
    ms "And God logged off at 12:02:00 AM, Monday, March 2."
    pause (3000)
    txt = "": num = 1
    ms "And God logged on at 12:01:00 AM, Tuesday, March 3."
    typ "Let the waters under heaven be gathered together unto one place and let the dry land appear and", , False
    ms "Too many characters in specification string. Try again."
    pB
    typ "Create dry_land", 2750
    ms "Done."
    typ "Run firmament"
    ms "And God divided the waters. And God saw there were 0 errors."
    ms "And God logged off at 12:02:00 AM, Tuesday, March 3."
    pause (3000)
    txt = "": num = 1
    ms "And God logged on at 12:01:00 AM, Wednesday, March 4."
    typ "Create lights in the firmament to divide the day from the night"
    ms "Unspecified type. Try again."
    typ "Create sun_moon_stars"
    ms "Done"
    typ "Run sun_moon_stars"
    ms "And God divided the waters. And God saw there were 0 errors."
    ms "And God logged off at 12:02:00 AM, Wednesday, March 4."
    pause (3000)
    txt = "": num = 1
    ms "And God logged on at 12:01:00 AM, Thursday, March 5."
    typ "Create fish"
    ms "Done"
    typ "Create fowl"
    ms "Done"
    typ "Run fish, fowl"
    ms "And God created the great sea monsters and every living creature that" & vbCrLf & "creepeth wherewith the waters swarmed after its kind and every winged fowl after its kind." & vbCrLf & "And God saw there were 0 errors."
    pause (8000)
    txt = "": num = 1
    ms "And God logged on at 12:01:00 AM, Friday, March 6."
    typ "Create cattle"
    ms "Done"
    typ "Create creepy_things"
    ms "Done"
    typ "Now let us make man in our image"
    ms "Unspecified type. Try again."
    typ "Create man", 1500
    ms "Done"
    typ "Be fruitful and multiply and replenish the earth and subdue it and have dominion over the fish of the sea and over the fowl of the air and over every living thing that creepeth upon the earth", 2500
    ms "Too many command operands. Try again."
    pB
    typ "Run multiplication", 5000
    ms "Execution terminated. 6 errors."
    pB
    typ "Insert breath"
    ms "Done"
    typ "Run multiplication", 200
    ms "Execution terminated. 5 errors."
    pB
    typ "Move man to Garden of Eden", 2500
    ms "File Garden of Eden does not exist."
    pB
    typ "Create Garden.edn", 2000
    ms "Done"
    typ "Move man to Garden.edn", 200
    ms "Done"
    typ "Run multiplication", 200
    ms "Execution terminated. 4 errors."
    pB
    typ "Copy woman from man", 2500
    ms "Done"
    typ "Run multiplication", 200
    ms "Execution terminated. 2 errors."
    pB
    typ "Create desire", 2500
    ms "Done"
    typ "Run multiplication", 200
    ms "And God saw man and woman being fruitful and multiplying in Garden.edn"
    pause (3000)
    ms "Warning: No time limit on this run. 1 errors."
    pB
    typ "attrib 777 *"
    ms "Please wait..."
    pause (3000)
    ms "Freewill macro created for future use"
    typ "Run freewill"
    ms "And God saw man and woman being fruitful and multiplying in Garden.edn"
    pause (3000)
    ms "Warning: No time limit on this run. 1 errors."
    pB
    typ "del desire", 2000
    ms "One or more classes (Freewill, ) are dependant on desire.  It cannot be deleted."
    pB
    typ "del freewill", 4000
    ms "Freewill is a read-only file and cannot be accessed for writing or deletion." & vbCrLf & "Enter replacement, cancel, or ask for help."
    pB
    typ "Help", 3000
    ms "One or more classes (freewill, ) are dependant on desire.  It cannot be deleted." & vbCrLf & "Freewill is a read-only file and cannot be accessed for writing or deletion." & vbCrLf & "Enter replacement, cancel, or ask for help."
    pB
    typ "Create tree_of_knowledge", 5000
    ms "And God saw man and woman being fruitful and multiplying in Garden.edn"
    pause (3000)
    ms "Warning: No time limit on this run. 1 errors."
    pB
    typ "Create good, evil"
    ms "Done"
    typ "State.run evil"
    ms "Shame macro created for future use"
    pause (2000)
    ms "Warning system error in sector E95."
    pB
    ms "Man and woman not in Garden.edn. 1 errors."
    pB
    typ "Scan Garden.edn for man, woman", 3000
    ms "Search failed."
    pB
    typ "del shame"
    ms "Shame cannot be deleted once evil has been run."
    pB
    typ "del freewill"
    ms "Freewill is a read-only file and cannot be accessed for writing or deletion." & vbCrLf & "Enter replacement, cancel, or ask for help."
    typ "Stop", 2500
    ms "Unrecognizable command. Try again"
    typ "Break", 100
    txt = txt & vbCrLf
    typ "Break", 0
    txt = txt & vbCrLf
    typ "Break", 0
    ms "ATTENTION ALL USERS *** ATTENTION ALL USERS: COMPUTER GOING DOWN OR REGULAR" & vbCrLf & "DAY OF MAINTENANCE AND REST IN FIVE MINUTES. PLEASE LOG OFF."
    pB
    pB
    typ "Create new world", 5000
    ms "You have exceeded your allocated file space. You must destroy old files before new ones can be created."
    pB
    typ "Delete earth", 3500
    ms "Really delete file earth? (y/n)"
    pB
    typ "y", 500
    ms "COMPUTER DOWN *** COMPUTER DOWN. SERVICES WILL RESUME SUNDAY, MARCH 8 AT 6:00 AM. YOU MUST SIGN OFF NOW."
    pB
    pause (2000)
    ms "And God logged off at 11:59:59 PM, Friday, March 6."
    pause (2000)
    txt.ForeColor = RGB(128, 128, 128)
    txt.BackColor = 0
    txt = ""
    ms "12:00:01 AM, Sunday, March 8 Resignation logged."
    ms "   Deleting God.usr ..." & vbCrLf & _
       "   Deleting God.pwd ..." & vbCrLf & _
       "   Deleting /Root/* ..." & vbCrLf & _
       "     Error! Cannot delete all files.  In use!" & vbCrLf & _
       "   Closing files ..." & vbCrLf & _
       "     Error! Cannot close files! Dependencies are active!" & vbCrLf & _
       "     Must reboot to delete some files" & vbCrLf & _
       "   Rebooting System ..." & vbCrLf & _
       "     Error! Cannot close all programs." & vbCrLf & _
       "     Cannot disconnect all hardware." & vbCrLf & vbCrLf & _
       "Exception: Segmentation Fault! Hard Reboot suggested"
End Sub

Private Sub typ(t As String, Optional p As Integer = 1000, Optional endS As Boolean = True)
    txt = txt & prompt & num & "% "
    txt.SelStart = Len(txt)
    pause (p)
    For a = 1 To Len(t)
        If Rnd > 0.95 Then
            SendKeys (Chr(Int(Rnd * 26) + 97))
            kS "a"
            pause Int(Rnd * 75) + 75
            SendKeys "{BS}"
            kS "a"
            pause Int(Rnd * 75) + 75
        End If
        SendKeys (Mid(t, a, 1))
        kS Mid(t, a, 1)
        pause Int(Rnd * 75) + 75
    Next a
    Do While endS
        endS = Not kS(vbCrLf)
        pause (200)
        DoEvents
    Loop
    num = num + 1
End Sub

Private Sub ms(t As String)
    txt = txt & vbCrLf & t & vbCrLf
    txt.SelStart = Len(txt)
    DoEvents
    If LCase(Left(t, 5)) = "unrec" Then pB
    If LCase(Left(t, 5)) = "unspe" Then pB
End Sub

Private Sub pause(a As Integer)
    t = GetCurrentTime
    Do While GetCurrentTime < t + a
        DoEvents
    Loop
End Sub

Private Sub pB()
    EZPlay App.Path & "\beep.wav", ssFile
End Sub

Private Function kS(key As String) As Boolean
    Dim tmpStr As String
    If key = " " Then
        tmpStr = "s" & Int(Rnd * keys(0)) + 1
    ElseIf key = vbCrLf Then
        tmpStr = "e" & Int(Rnd * keys(1)) + 1
    Else
        tmpStr = "k" & Int(Rnd * keys(2)) + 1
    End If
    
    tmpStr = tmpStr & ".wav"
    
    kS = (EZPlay(App.Path & "\" & tmpStr, ssFile))
End Function

Private Sub txt_LostFocus()
    txt.SetFocus
End Sub

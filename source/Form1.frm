VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "toolBox"
   ClientHeight    =   2610
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   7665
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "always on top"
      Height          =   255
      Left            =   6120
      TabIndex        =   24
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "local admin"
      Height          =   735
      Index           =   22
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   21
      Left            =   2640
      Picture         =   "Form1.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   20
      Left            =   960
      Picture         =   "Form1.frx":1E84
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   19
      Left            =   6840
      Picture         =   "Form1.frx":2EC6
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   18
      Left            =   6000
      Picture         =   "Form1.frx":3F08
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   12
      Left            =   120
      Picture         =   "Form1.frx":4F4A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   6
      Left            =   960
      Picture         =   "Form1.frx":5F8C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":6FCE
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   17
      Left            =   4320
      Picture         =   "Form1.frx":8010
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   16
      Left            =   1800
      Picture         =   "Form1.frx":9052
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   15
      Left            =   2640
      Picture         =   "Form1.frx":A094
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   14
      Left            =   960
      Picture         =   "Form1.frx":B0D6
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   13
      Left            =   120
      Picture         =   "Form1.frx":C118
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   11
      Left            =   3480
      Picture         =   "Form1.frx":D15A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "printer"
      Height          =   735
      Index           =   10
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   9
      Left            =   3480
      Picture         =   "Form1.frx":E19C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   8
      Left            =   5160
      Picture         =   "Form1.frx":F1DE
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "folder options"
      Height          =   735
      Index           =   7
      Left            =   5160
      Picture         =   "Form1.frx":10220
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   5
      Left            =   1800
      Picture         =   "Form1.frx":11262
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   4
      Left            =   1800
      Picture         =   "Form1.frx":122A4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000005&
      Caption         =   "local security policy"
      Height          =   735
      Index           =   3
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   2
      Left            =   2640
      Picture         =   "Form1.frx":132E6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "data source"
      Height          =   735
      Index           =   1
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   22
      Top             =   1920
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written by KramWell.com - 29/NOV/2015
'Easy tool to quickly access the most useful windows programs all from one location.

Option Explicit

Private Declare Function IsUserAdmin Lib "Shell32" Alias "#680" () As Boolean

'-=-==-=- always on top
      Private Const SWP_NOMOVE = 2
      Private Const SWP_NOSIZE = 1
      Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
      Private Const HWND_TOPMOST = -1
      Private Const HWND_NOTOPMOST = -2

      Private Declare Function SetWindowPos Lib "user32" _
            (ByVal hwnd As Long, _
            ByVal hWndInsertAfter As Long, _
            ByVal x As Long, _
            ByVal y As Long, _
            ByVal cx As Long, _
            ByVal cy As Long, _
            ByVal wFlags As Long) As Long
            
      Private Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
         As Long

         If Topmost = True Then 'Make the window topmost
            SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, _
               0, FLAGS)
         Else
            SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, _
               0, 0, FLAGS)
            SetTopMostWindow = False
         End If
      End Function

'-=-==-=- always on top

Private Sub Check2_Click()
         Dim lR As Long
If Check2.Value = 1 Then 'checked
         lR = SetTopMostWindow(Form1.hwnd, True)
         Else
         lR = SetTopMostWindow(Form1.hwnd, False)
End If
End Sub

Private Sub error_Cancel()

If Err.Number <> 0 Then
MsgBox Err.Description, vbCritical, "Error " & Err.Number
End If

End Sub

Private Sub cmdRun_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo errorrepair

'admin tools start

'computer management
        If Index = 0 Then
            Label1.Caption = "computer management"
'data source
        ElseIf Index = 1 Then
            Label1.Caption = "data source"
'event viewer
        ElseIf Index = 2 Then
            Label1.Caption = "event viewer"
'local security policy
        ElseIf Index = 3 Then
            Label1.Caption = "local security policy"
'services
        ElseIf Index = 4 Then
            Label1.Caption = "services"

'Control Panel start
    
    'add/remove
        ElseIf Index = 5 Then
            Label1.Caption = "add/remove"
    'display properties
        ElseIf Index = 6 Then
            Label1.Caption = "display properties"
    'folder options
        ElseIf Index = 7 Then
            Label1.Caption = "folder options"
    'internet options
        ElseIf Index = 8 Then
            Label1.Caption = "internet options"
    'network connections
        ElseIf Index = 9 Then
            Label1.Caption = "network connections"
    'printer
        ElseIf Index = 10 Then
            Label1.Caption = "printer"
    'system properties
        ElseIf Index = 11 Then
            Label1.Caption = "system properties"
    'user/accounts
        ElseIf Index = 12 Then
            Label1.Caption = "user/accounts"

'Control Panel end
    
'extras start

    'cmd prompt
        ElseIf Index = 13 Then
            Label1.Caption = "cmd prompt"
    'device manager
        ElseIf Index = 14 Then
            Label1.Caption = "device manager"
    'nslookup
        ElseIf Index = 15 Then
            Label1.Caption = "nslookup"
    'registry editor
        ElseIf Index = 16 Then
            Label1.Caption = "registry editor"
    'task manager
        ElseIf Index = 17 Then
            Label1.Caption = "task manager"
    '--------------
    'disk defrag
        ElseIf Index = 18 Then
            Label1.Caption = "disk defrag"
    'disk cleanup
        ElseIf Index = 19 Then
            Label1.Caption = "disk cleanup"
    '-------------
    'remote desktop
        ElseIf Index = 20 Then
            Label1.Caption = "remote desktop"
    'shared folders
        ElseIf Index = 21 Then
            Label1.Caption = "shared folders"
'extras ends

    'local admin
        ElseIf Index = 22 Then
            Label1.Caption = "local admin"

End If

errorrepair:
error_Cancel

End Sub

Private Sub cmdRun_Click(Index As Integer)
Dim taskid As Long

On Error GoTo errorrepair

'admin tools start

'computer management
        If Index = 0 Then
            taskid = Shell("mmc.exe compmgmt.msc /s", vbNormalFocus)
'data source
        ElseIf Index = 1 Then
            taskid = Shell("odbcad32.exe", vbNormalFocus)
'event viewer
        ElseIf Index = 2 Then
            taskid = Shell("mmc.exe eventvwr.msc /s", vbNormalFocus)
'local security policy
        ElseIf Index = 3 Then
            taskid = Shell("mmc.exe secpol.msc /s", vbNormalFocus)
'services
        ElseIf Index = 4 Then
            taskid = Shell("mmc.exe services.msc /s", vbNormalFocus)

'Control Panel start
    
    'add/remove
        ElseIf Index = 5 Then
                taskid = Shell("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl", vbNormalFocus)
    'display properties
        ElseIf Index = 6 Then
                taskid = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl", vbNormalFocus)
    'folder options
        ElseIf Index = 7 Then
                taskid = Shell("rundll32.exe shell32.dll,Options_RunDLL 0", vbNormalFocus)
    'internet options
        ElseIf Index = 8 Then
                taskid = Shell("rundll32.exe shell32.dll,Control_RunDLL Inetcpl.cpl", vbNormalFocus)
              'taskid = Shell("javaw.exe,@1", vbNormalFocus)
    'network connections
        ElseIf Index = 9 Then
                taskid = Shell("rundll32.exe shell32.dll,Control_RunDLL ncpa.cpl", vbNormalFocus)
    'printer
        ElseIf Index = 10 Then
                taskid = Shell("explorer.exe ::{2227A280-3AEA-1069-A2DE-08002B30309D}", vbNormalFocus)
    'system properties
        ElseIf Index = 11 Then
                taskid = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl", vbNormalFocus)
    'user/accounts
        ElseIf Index = 12 Then
                taskid = Shell("rundll32.exe shell32.dll,Control_RunDLL nusrmgr.cpl", vbNormalFocus)

'Control Panel end
    
'extras start

    'cmd prompt
        ElseIf Index = 13 Then
            taskid = Shell("cmd.exe", vbNormalFocus)
    'device manager
        ElseIf Index = 14 Then
                taskid = Shell("mmc.exe devmgmt.msc", vbNormalFocus)
    'nslookup
        ElseIf Index = 15 Then
                taskid = Shell("nslookup.exe", vbNormalFocus)
    'registry editor
        ElseIf Index = 16 Then
                taskid = Shell("regedt32.exe", vbNormalFocus)
    'task manager
        ElseIf Index = 17 Then
                taskid = Shell("taskmgr.exe", vbNormalFocus)
    '--------------
    'disk defrag
        ElseIf Index = 18 Then
                taskid = Shell("mmc.exe dfrg.msc", vbNormalFocus)
    'disk cleanup
        ElseIf Index = 19 Then
                taskid = Shell("cleanmgr.exe", vbNormalFocus)
    '-------------
    'remote desktop
        ElseIf Index = 20 Then
                taskid = Shell("mstsc.exe", vbNormalFocus)
    'shared folders
        ElseIf Index = 21 Then
                taskid = Shell("mmc.exe fsmgmt.msc", vbNormalFocus)

'extras ends

    'local admin
        ElseIf Index = 22 Then
                taskid = Shell("mmc.exe lusrmgr.msc", vbNormalFocus)


End If

errorrepair:
error_Cancel

Label1.Caption = Index
End Sub

Private Sub Form_Load()

Dim UserName As String
Dim UserDomain As String
UserName = Environ("USERNAME")
UserDomain = Environ("USERDOMAIN")

If IsUserAdmin() = 0 Then
  MsgBox "Please be aware you are NOT ADMIN"
  'Unload Me
End If

Form1.Caption = "toolBox - " & UserDomain & "\" & UserName

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.Caption = ""
End Sub


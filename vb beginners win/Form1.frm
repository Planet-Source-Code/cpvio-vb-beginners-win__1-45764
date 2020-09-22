VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command13 
      Caption         =   "recyle bin"
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "yea"
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "right click clock"
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "clear text"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "change start caption"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "hide clock"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "show clock "
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "hide bar"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "show bar"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "left click start"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "right click start"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "show start"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "hide start"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFFF00&
      Height          =   255
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   135
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF00FF&
      Height          =   495
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   735
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   975
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1215
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "enter sbutton caption"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim shelltraywnd As Long, button As Long, rebarwindow As Long, traynotifywnd As Long, trayclockwclass As Long
Dim explorewclass As Long, shelldlldefview As Long, duiviewwndclassname As Long
Dim directuihwnd As Long, ctrlnotifysink As Long, syslistview As Long

Private Sub Command1_Click()
'this hides the start button
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
button = FindWindowEx(shelltraywnd, 0&, "button", vbNullString)
Call ShowWindow(button, SW_HIDE)
End Sub

Private Sub Command10_Click()
'this clears the text nothing
Text1.Text = ""
End Sub



Private Sub Command11_Click()
'this finds the right mouse click on the system clock and clicks it
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
traynotifywnd = FindWindowEx(shelltraywnd, 0&, "traynotifywnd", vbNullString)
trayclockwclass = FindWindowEx(traynotifywnd, 0&, "trayclockwclass", vbNullString)
Call SendMessageLong(trayclockwclass, WM_RBUTTONDOWN, 0&, 0&)
Call SendMessageLong(trayclockwclass, WM_RBUTTONUP, 0&, 0&)

End Sub

Private Sub Command12_Click()
'ummm you can quess
Text1.Text = "***PSC Rockz 2003***"
End Sub

Private Sub Command13_Click()
'finds the window handle of the recycle bin
Dim TheWin As Long
TheWin = find_syslistview()
 If TheWin <> 0 Then
' What to do if window is there
Text1.Text = TheWin
 End If
End Sub

Private Sub Command2_Click()
'this shows  the start button
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
button = FindWindowEx(shelltraywnd, 0&, "button", vbNullString)
Call ShowWindow(button, SW_SHOW)
End Sub

Private Sub Command3_Click()
'this right clicks the start button to view the menu
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
button = FindWindowEx(shelltraywnd, 0&, "button", vbNullString)
Call SendMessageLong(button, WM_RBUTTONDOWN, 0&, 0&)
Call SendMessageLong(button, WM_RBUTTONUP, 0&, 0&)
End Sub

Private Sub Command4_Click()
'this left clicks the start button to view the menu
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
button = FindWindowEx(shelltraywnd, 0&, "button", vbNullString)
Call SendMessageLong(button, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(button, WM_LBUTTONUP, 0&, 0&)
End Sub

Private Sub Command5_Click()
'this shows the running apps in system bar
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
rebarwindow = FindWindowEx(shelltraywnd, 0&, "rebarwindow32", vbNullString)
Call ShowWindow(rebarwindow, SW_SHOW)
End Sub

Private Sub Command6_Click()
'this hides the running apps in system bar
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
rebarwindow = FindWindowEx(shelltraywnd, 0&, "rebarwindow32", vbNullString)
Call ShowWindow(rebarwindow, SW_HIDE)
End Sub

Private Sub Command7_Click()
'this shows the clock in system tray
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
traynotifywnd = FindWindowEx(shelltraywnd, 0&, "traynotifywnd", vbNullString)
trayclockwclass = FindWindowEx(traynotifywnd, 0&, "trayclockwclass", vbNullString)
Call ShowWindow(trayclockwclass, SW_SHOW)
End Sub

Private Sub Command8_Click()
'this hides the clock in system tray
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
traynotifywnd = FindWindowEx(shelltraywnd, 0&, "traynotifywnd", vbNullString)
trayclockwclass = FindWindowEx(traynotifywnd, 0&, "trayclockwclass", vbNullString)
Call ShowWindow(trayclockwclass, SW_HIDE)
End Sub

Private Sub Command9_Click()
'this finds the start buttons caption and changes it to whats in the text1.text box
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
button = FindWindowEx(shelltraywnd, 0&, "button", vbNullString)
Call SendMessageByString(button, WM_SETTEXT, 0&, Text1.Text)
End Sub
Public Function find_syslistview() As Long
' If this function finds the window, it will return it's
' handle. If it doesn't find it, it will return 0.
Dim explorewclass As Long, shelldlldefview As Long, duiviewwndclassname As Long
Dim directuihwnd As Long, ctrlnotifysink As Long, syslistview As Long
explorewclass = FindWindow("explorewclass", vbNullString)
shelldlldefview = FindWindowEx(explorewclass, 0&, "shelldll_defview", vbNullString)
duiviewwndclassname = FindWindowEx(shelldlldefview, 0&, "duiviewwndclassname", vbNullString)
directuihwnd = FindWindowEx(duiviewwndclassname, 0&, "directuihwnd", vbNullString)
ctrlnotifysink = FindWindowEx(directuihwnd, 0&, "ctrlnotifysink", vbNullString)
syslistview = FindWindowEx(ctrlnotifysink, 0&, "syslistview32", vbNullString)
Dim Winkid1 As Long, FindOtherWin As Long
FindOtherWin = GetWindow(syslistview, GW_HWNDFIRST)
Do While FindOtherWin <> 0
       DoEvents
       Winkid1 = FindWindowEx(FindOtherWin, 0&, "sysheader32", vbNullString)
       If (Winkid1 <> 0) Then
              find_syslistview = FindOtherWin
              Exit Function
       End If
       FindOtherWin = GetWindow(FindOtherWin, GW_HWNDNEXT)
Loop
find_syslistview = 0
' example on how to use:
' Dim TheWin As Long
' TheWin = find_syslistview()
' If TheWin <> 0 Then
' What to do if window is there
' End If
End Function

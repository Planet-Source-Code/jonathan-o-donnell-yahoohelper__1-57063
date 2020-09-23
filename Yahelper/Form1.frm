VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "YaHelper"
   ClientHeight    =   7230
   ClientLeft      =   2055
   ClientTop       =   2835
   ClientWidth     =   11190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   11190
   Begin VB.CommandButton Command7 
      Caption         =   "Exit"
      Height          =   375
      Left            =   9720
      TabIndex        =   9
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Go Forward"
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Go Back"
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "GeoCities"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Yahoo Games"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yahoo Mail"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   6480
      Width           =   10935
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   6015
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   10695
      ExtentX         =   18865
      ExtentY         =   10610
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Main Screen"
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
URL = "http://chat.yahoo.com"
wb.Navigate URL
End Sub

Private Sub Command1_Click()
On Error Resume Next
URL = "http://us.f210.mail.yahoo.com"
wb.Navigate URL
End Sub

Private Sub Command2_Click()
On Error Resume Next
URL = "http://games.yahoo.com"
wb.Navigate URL
End Sub

Private Sub Command3_Click()
On Error Resume Next
URL = "http://geocities.yahoo.com/home/"
wb.Navigate URL
End Sub

Private Sub Command4_Click()
On Error Resume Next
wb.GoBack
End Sub

Private Sub Command5_Click()
On Error Resume Next
wb.GoForward
End Sub

Private Sub Command6_Click()
On Error Resume Next
wb.Refresh
End Sub

Private Sub Command7_Click()
On Error Resume Next
MsgBox "Yahoo Helper By JizZy Pe@ce", vbInformation, "Exit"
End
End Sub



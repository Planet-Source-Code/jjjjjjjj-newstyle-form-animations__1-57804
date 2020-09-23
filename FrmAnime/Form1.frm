VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Select the AnimationType"
   ClientHeight    =   1665
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstAnime 
      Height          =   1500
      ItemData        =   "Form1.frx":0000
      Left            =   3480
      List            =   "Form1.frx":0013
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Animated  Form"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select the 'Animation Speed' in the code.Now it is ""Medium""."
      Height          =   600
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLoad_Click()
    frmAnime.Show
End Sub

Private Sub Form_Load()
    lstAnime.Selected(0) = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Is it Satisfactory?", vbQuestion + vbYesNo, "Please tell Me") = vbYes Then
        MsgBox "(  PLEASE 'RATE' THIS CODE  ).I want to know how do you rate this code.The site address will be copied to your clipboard", vbInformation, "ThankYou"
    Else
        MsgBox "( PLEASE GIVE FEEDBACK ) to improve this code.The site address will be copied to your clipboard", vbInformation, "Please Give FeedBack"
    End If
    Clipboard.SetText ("http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57804&lngWId=1")
End Sub

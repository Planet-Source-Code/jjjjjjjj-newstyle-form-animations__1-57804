VERSION 5.00
Begin VB.Form frmAnime 
   Caption         =   "Animated  Form"
   ClientHeight    =   4740
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   8580
   Icon            =   "frmAnime.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4740
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUnload 
      Caption         =   "Unload"
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Please Set the 'Animated Form' as the 'StartUp Object'  before compiling."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Width           =   7695
   End
   Begin VB.Label Label1 
      Caption         =   "You must Set a  definite 'AnimationType' and exicute it after compiling to get the actual result."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   7680
   End
End
Attribute VB_Name = "frmAnime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdUnload_Click()
    Unload Me
End Sub

'Change the Animation Speed
Private Sub Form_Load()
    AnimateForm Me, aLoad, Form1.lstAnime.ListIndex, aMedium
End Sub

'Change the Animation Speed
Private Sub Form_Unload(Cancel As Integer)
    AnimateForm Me, aUnload, Form1.lstAnime.ListIndex, aMedium
End Sub


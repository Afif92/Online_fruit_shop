VERSION 5.00
Begin VB.Form welcome 
   Caption         =   "Welcme"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8460
   LinkTopic       =   "welcome"
   ScaleHeight     =   7500
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   6135
      Left            =   0
      Picture         =   "welcome.frx":0000
      ScaleHeight     =   6075
      ScaleWidth      =   8355
      TabIndex        =   0
      Top             =   0
      Width           =   8415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Welcome To Online Fruits Shop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   6240
      Width           =   8295
   End
End
Attribute VB_Name = "welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    index.Show
    Me.Hide
End Sub

VERSION 5.00
Begin VB.Form grape 
   Caption         =   "Grape"
   ClientHeight    =   8355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7290
   LinkTopic       =   "yellow"
   ScaleHeight     =   8355
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3555
      ScaleWidth      =   7035
      TabIndex        =   8
      Top             =   120
      Width           =   7095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   735
      Left            =   5280
      TabIndex        =   7
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   735
      Left            =   5280
      TabIndex        =   6
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Purchase"
      Height          =   735
      Left            =   5280
      TabIndex        =   5
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   6960
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   6000
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   5040
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Text            =   "Select to see"
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Colour"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3960
      Width           =   1695
   End
End
Attribute VB_Name = "grape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
       purchase.Show
       
End Sub

Private Sub Command2_Click()
    index.Show
    Me.Hide
    
End Sub

Private Sub Command3_Click()
    End
    
End Sub

Private Sub Form_Load()
     Combo1.AddItem "1 kg red Grape"
     Combo1.AddItem "1 kg green Grape"
End Sub
Private Sub combo1_Click()
    If Combo1.ListIndex = 0 Then
        Picture1.Picture = LoadPicture("D:\Online fruit shop\pic\grapes.jpg")
        Text1.Text = "300"
        Text2.Text = "Grape"
        Text3.Text = "1 kg red Grape"
    End If
    If Combo1.ListIndex = 1 Then
        Picture1.Picture = LoadPicture("D:\Online fruit shop\pic\1kg green grape.jpg")
        Text1.Text = "260"
        Text2.Text = "Grape"
        Text3.Text = "1 kg green Grape"
    End If
    
End Sub

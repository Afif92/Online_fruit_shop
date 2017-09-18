VERSION 5.00
Begin VB.Form banana 
   Caption         =   "Banana"
   ClientHeight    =   8355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7290
   LinkTopic       =   "blue"
   ScaleHeight     =   8355
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   0
      ScaleHeight     =   4035
      ScaleWidth      =   7155
      TabIndex        =   8
      Top             =   0
      Width           =   7215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   735
      Left            =   4920
      TabIndex        =   7
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   735
      Left            =   4920
      TabIndex        =   6
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Purchase"
      Height          =   735
      Left            =   4920
      TabIndex        =   5
      Top             =   5040
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
      Width           =   2055
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
      Width           =   2055
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
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Text            =   "Select to see"
      Top             =   4320
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
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   1695
   End
End
Attribute VB_Name = "banana"
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
     Combo1.AddItem "Single banana"
     Combo1.AddItem "A dozen of banana"
End Sub
Private Sub combo1_Click()
    If Combo1.ListIndex = 0 Then
        Picture1.Picture = LoadPicture("D:\Online fruit shop\pic\banana.jpg")
        Text1.Text = "5"
        Text2.Text = "Banana"
        Text3.Text = "Single"
    End If
    If Combo1.ListIndex = 1 Then
        Picture1.Picture = LoadPicture("D:\Online fruit shop\pic\dozen banana.jpg")
        Text1.Text = "60"
        Text2.Text = "Banana"
        Text3.Text = "A dozen of banana"
    End If
    
End Sub

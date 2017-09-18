VERSION 5.00
Begin VB.Form index 
   Caption         =   "Index"
   ClientHeight    =   6990
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11955
   LinkTopic       =   "index"
   ScaleHeight     =   6990
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Back"
      Height          =   735
      Left            =   8400
      TabIndex        =   16
      Top             =   3960
      Width           =   3255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Exit"
      Height          =   735
      Left            =   8280
      TabIndex        =   15
      Top             =   5400
      Width           =   3495
   End
   Begin VB.Frame Frame5 
      Caption         =   "Apple"
      Height          =   3015
      Left            =   4200
      TabIndex        =   12
      Top             =   3720
      Width           =   3615
      Begin VB.CommandButton Command5 
         Caption         =   "Show Details"
         Height          =   615
         Left            =   0
         TabIndex        =   14
         Top             =   2400
         Width           =   3615
      End
      Begin VB.PictureBox Picture5 
         Height          =   2055
         Left            =   0
         Picture         =   "index.frx":0000
         ScaleHeight     =   1995
         ScaleWidth      =   3555
         TabIndex        =   13
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Grape"
      Height          =   3015
      Left            =   360
      TabIndex        =   9
      Top             =   3720
      Width           =   3615
      Begin VB.PictureBox Picture4 
         Height          =   2055
         Left            =   0
         Picture         =   "index.frx":1C87
         ScaleHeight     =   1995
         ScaleWidth      =   3555
         TabIndex        =   11
         Top             =   360
         Width           =   3615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Show Details"
         Height          =   615
         Left            =   0
         TabIndex        =   10
         Top             =   2400
         Width           =   3615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Guaba"
      Height          =   3015
      Left            =   8040
      TabIndex        =   6
      Top             =   120
      Width           =   3615
      Begin VB.PictureBox Picture3 
         Height          =   2055
         Left            =   0
         Picture         =   "index.frx":2D6A
         ScaleHeight     =   1995
         ScaleWidth      =   3555
         TabIndex        =   8
         Top             =   360
         Width           =   3615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Show Details"
         Height          =   615
         Left            =   0
         TabIndex        =   7
         Top             =   2400
         Width           =   3615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Banana"
      Height          =   3015
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   3615
      Begin VB.PictureBox Picture2 
         Height          =   2055
         Left            =   0
         Picture         =   "index.frx":3C00
         ScaleHeight     =   1995
         ScaleWidth      =   3555
         TabIndex        =   5
         Top             =   360
         Width           =   3615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Show Details"
         Height          =   615
         Left            =   0
         TabIndex        =   4
         Top             =   2400
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Orrange"
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton Command1 
         Caption         =   "Show Details"
         Height          =   615
         Left            =   0
         TabIndex        =   2
         Top             =   2400
         Width           =   3615
      End
      Begin VB.PictureBox Picture1 
         Height          =   2055
         Left            =   0
         Picture         =   "index.frx":4BF0
         ScaleHeight     =   1995
         ScaleWidth      =   3555
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
   End
End
Attribute VB_Name = "index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    orrange.Show
    Me.Hide
    
End Sub

Private Sub Command2_Click()
banana.Show
Me.Hide

End Sub

Private Sub Command3_Click()
    guaba.Show
    Me.Hide
    
End Sub

Private Sub Command4_Click()
    grape.Show
    Me.Hide
    
End Sub

Private Sub Command5_Click()
    apple.Show
    Me.Hide
    
End Sub

Private Sub Command7_Click()
     End
     
End Sub

Private Sub Command8_Click()
    welcome.Show
    Me.Hide
    
End Sub

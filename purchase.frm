VERSION 5.00
Begin VB.Form purchase 
   Caption         =   "Purchase"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13065
   LinkTopic       =   "purchase"
   ScaleHeight     =   8385
   ScaleWidth      =   13065
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox mb 
      Height          =   495
      Left            =   11640
      TabIndex        =   55
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox yb 
      Height          =   495
      Left            =   9120
      TabIndex        =   54
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox ys 
      Height          =   495
      Left            =   6480
      TabIndex        =   53
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox wb 
      Height          =   495
      Left            =   3960
      TabIndex        =   52
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox ws 
      Height          =   495
      Left            =   1320
      TabIndex        =   51
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox ms 
      Height          =   495
      Left            =   11640
      TabIndex        =   50
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox bb 
      Height          =   495
      Left            =   9120
      TabIndex        =   49
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox bs 
      Height          =   495
      Left            =   6480
      TabIndex        =   48
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox rb 
      Height          =   495
      Left            =   3960
      TabIndex        =   47
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame Frame9 
      Caption         =   "Apple"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   10560
      TabIndex        =   38
      Top             =   240
      Width           =   2415
      Begin VB.CheckBox mis 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   255
      End
      Begin VB.Frame Frame12 
         Caption         =   "Select"
         Height          =   1575
         Left            =   2640
         TabIndex        =   39
         Top             =   0
         Width           =   2415
         Begin VB.OptionButton Option23 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1440
            TabIndex        =   42
            Top             =   960
            Width           =   735
         End
         Begin VB.OptionButton Option24 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   41
            Top             =   960
            Width           =   615
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Check1"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label12 
            Caption         =   "White rose bouquet"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   43
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Label Label20 
         Caption         =   "Quantity"
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "single apple"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   45
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Coustomer info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   240
      TabIndex        =   30
      Top             =   4680
      Width           =   12255
      Begin VB.CommandButton Command2 
         Caption         =   "Confirm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   8280
         TabIndex        =   37
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox ad 
         Height          =   1095
         Left            =   2040
         TabIndex        =   36
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox ph 
         Height          =   495
         Left            =   7800
         TabIndex        =   34
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox na 
         Height          =   405
         Left            =   2040
         TabIndex        =   32
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label26 
         Caption         =   "Select items ,Fill up this form and press Confirm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   66
         Top             =   240
         Width           =   6255
      End
      Begin VB.Label Label16 
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   35
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6240
         TabIndex        =   33
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   31
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10080
      MaskColor       =   &H00808080&
      TabIndex        =   29
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Frame Frame10 
      Caption         =   "Apple"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   10560
      TabIndex        =   24
      Top             =   2040
      Width           =   2415
      Begin VB.CheckBox mib 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label25 
         Caption         =   "Quantity"
         Height          =   375
         Left            =   120
         TabIndex        =   65
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "1 kg apple"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Grape"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8040
      TabIndex        =   21
      Top             =   2040
      Width           =   2415
      Begin VB.CheckBox yeb 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label24 
         Caption         =   "Quantity"
         Height          =   375
         Left            =   120
         TabIndex        =   64
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "1 kg green grape"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Grape"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5400
      TabIndex        =   18
      Top             =   2040
      Width           =   2415
      Begin VB.CheckBox yes 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label23 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "1 kg red grape"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Guaba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2880
      TabIndex        =   15
      Top             =   2040
      Width           =   2415
      Begin VB.CheckBox whb 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label22 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "1 kg guaba"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Guaba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   2415
      Begin VB.CheckBox whs 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "Quantity"
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "single guaba"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Banana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8040
      TabIndex        =   9
      Top             =   240
      Width           =   2415
      Begin VB.CheckBox blb 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label19 
         Caption         =   "Quantity"
         Height          =   375
         Left            =   120
         TabIndex        =   59
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "1 kg banana"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Banana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5400
      TabIndex        =   6
      Top             =   240
      Width           =   2415
      Begin VB.CheckBox bls 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "Quantity"
         Height          =   375
         Left            =   120
         TabIndex        =   58
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "single banana"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Orrange"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   2415
      Begin VB.CheckBox reb 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label17 
         Caption         =   "Quantity"
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "1 kg orrange"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Orrange"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2415
      Begin VB.TextBox rs 
         Height          =   495
         Left            =   1080
         TabIndex        =   46
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox res 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "Quantity"
         Height          =   375
         Left            =   120
         TabIndex        =   56
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "single Orrange"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Label Label13 
      Caption         =   "Total Taka You have to pay"
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
      Left            =   360
      TabIndex        =   27
      Top             =   3840
      Width           =   5535
   End
End
Attribute VB_Name = "purchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

           Dim rsprice As Integer
           Dim rbprice As Integer
           Dim bsprice As Integer
           Dim bbprice As Integer
           Dim wsprice As Integer
           Dim wbprice As Integer
           Dim ysprice As Integer
           Dim ybprice As Integer
           Dim msprice As Integer
           Dim mbprice As Integer




           If res.Value = 0 Then
             rsprice = 0
             ElseIf res.Value = 1 Then
             rsprice = (15 * Val(rs.Text))
             End If
             
          If reb.Value = 0 Then
             rbprice = 0
             ElseIf reb.Value = 1 Then
             rbprice = (90 * Val(rb.Text))
             End If
              
           If bls.Value = 0 Then
             bsprice = 0
             ElseIf bls.Value = 1 Then
             bsprice = (5 * Val(bs.Text))
             End If
              
             
           If blb.Value = 0 Then
             bbprice = 0
             ElseIf blb.Value = 1 Then
             bbprice = (60 * Val(bb.Text))
             End If
           
           If mis.Value = 0 Then
             msprice = 0
             ElseIf mis.Value = 1 Then
             msprice = (10 * Val(ms.Text))
             End If
             
          If mib.Value = 0 Then
             mbprice = 0
             ElseIf mib.Value = 1 Then
             mbprice = (100 * Val(mb.Text))
             End If
              
           If whs.Value = 0 Then
             wsprice = 0
             ElseIf whs.Value = 1 Then
             wsprice = (300 * Val(ws.Text))
             End If
              
             
           If whb.Value = 0 Then
             wbprice = 0
             ElseIf whb.Value = 1 Then
             wbprice = (260 * Val(wb.Text))
             End If
             
           If yes.Value = 0 Then
             ysprice = 0
             ElseIf yes.Value = 1 Then
             ysprice = (15 * Val(ys.Text))
             End If
              
             
           If yeb.Value = 0 Then
             ybprice = 0
             ElseIf yeb.Value = 1 Then
             ybprice = (90 * Val(yb.Text))
             End If
            
             
Dim total As Integer
   total = Val(rsprice) + Val(rbprice) + Val(bsprice) + Val(bbprice) + Val(wsprice) + Val(wbprice) + Val(ysprice) + Val(ybprice) + Val(msprice) + Val(mbprice)
   Text1.Text = Val(total)
End Sub

Private Sub Command2_Click()
       
       
       MsgBox "Hay " + na.Text + "!!!!!!! " + Chr(13) + "Address- " + ad.Text + Chr(13) + "Mobile - " + ph.Text + Chr(13) + "Your order has confirmed and You will get your product within 1 hour" + Chr(13) + "You have to pay after get the product!!!!!!!!!"

End Sub

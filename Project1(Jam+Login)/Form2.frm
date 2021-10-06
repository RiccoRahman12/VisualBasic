VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   0
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2880
         TabIndex        =   7
         Top             =   3480
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   6
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   2160
         TabIndex        =   5
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   2160
         TabIndex        =   4
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Bernard MT Condensed"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Username :"
         BeginProperty Font 
            Name            =   "Bernard MT Condensed"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   0
         X2              =   5040
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1320
         TabIndex        =   1
         Top             =   120
         Width           =   2535
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "admin" And Text2 = "admin" Then
Form1.Show 'Perintah Menampilkan Form 2
Form2.Visible = False 'Menyembunyikan Form 1
Unload Me 'Menutup Form 1
Else
MsgBox "User Name atau Password yang Anda Masukkan salah" _
& vbNewLine & "Silahkan Coba lagi !!", vbCritical, "Warning!!"
Text1 = ""
Text2 = ""
Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
pesan = MsgBox("Anda Yakin Mau Keluar ??", vbQuestion + vbYesNo, "Question")
If pesan = vbYes Then
End
Else
Form2.SetFocus
End If
End Sub

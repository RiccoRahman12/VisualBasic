VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   480
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtth 
      Height          =   285
      Left            =   8160
      TabIndex        =   22
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox txtjb 
      Height          =   285
      Left            =   8160
      TabIndex        =   21
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox txtht 
      Height          =   285
      Left            =   8160
      TabIndex        =   20
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Jenis Pelayanan"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5760
      TabIndex        =   14
      Top             =   1440
      Width           =   3735
      Begin VB.OptionButton rdoe 
         Caption         =   "Executive"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   16
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton rdob 
         Caption         =   "Businese"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.ComboBox cmbj 
      Height          =   315
      Left            =   2760
      TabIndex        =   13
      Top             =   3960
      Width           =   2535
   End
   Begin VB.ComboBox cmbkk 
      Height          =   315
      Left            =   2760
      TabIndex        =   12
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox txtnke 
      Height          =   285
      Left            =   2760
      TabIndex        =   11
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox txtnp 
      Height          =   285
      Left            =   2760
      TabIndex        =   10
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtnktp 
      Height          =   285
      Left            =   2760
      TabIndex        =   9
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox txtnko 
      Height          =   285
      Left            =   2760
      TabIndex        =   8
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proses"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      TabIndex        =   6
      Top             =   4920
      Width           =   8655
      Begin VB.CommandButton Command2 
         Caption         =   "Keluar"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton btninput 
         Caption         =   "Input Data"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Travel"
      Height          =   615
      Left            =   3960
      TabIndex        =   24
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label10 
      Caption         =   "Total Harga"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   19
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label9 
      Caption         =   "Jumlah Beli"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   18
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "Harga Tiket"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   17
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Jurusan"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Nama Kereta"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Nama Pembeli"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "No KTP"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Kode"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   11160
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label5 
      Caption         =   "Kode Kereta"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
cmbkk.AddItem ("ARB")
cmbkk.AddItem ("JES")
cmbkk.AddItem ("KKD")
cmbj.AddItem ("cikampek-tuparev")
cmbj.AddItem ("johar-bypass")
cmbj.AddItem ("klari-badami")
tidakaktif
End Sub

Private Sub cmbkk_Click()
Select Case (cmbkk.Text)
Case "ARB"
txtnke.Text = "ARGO BISNIS"
Case "JES"
txtnke.Text = "JOHAR EXPRES"
Case "KKD"
txtnke.Text = "KERETA KUDA"
End Select
End Sub

Private Sub rdob_Click()
If cmbj.Text = "cikampek-tuparev" Then
txtht.Text = 20000
ElseIf cmbj.Text = "klari-badami" Then
txtht.Text = 15000
Else
txtht.Text = 10000
End If
End Sub

Private Sub rdoe_Click()
If cmbj.Text = "cikampek-tuparev" Then
txtht.Text = 30000
ElseIf cmbj.Text = "klari-badami" Then
txtht.Text = 25000
Else
txtht.Text = 20000
End If
End Sub

Private Sub txtjb_Change()
txtth.Text = Val(txtjb.Text) * Val(txtht.Text)
End Sub

Private Sub btninput_Click()
If btninput.Caption = "Input Data" Then
aktif
btninput.Caption = "Bersih"
Else
bersih
tidakaktif
btninput.Caption = "Input Data"
End If
End Sub

Sub aktif()
txtnko.Enabled = True
txtnp.Enabled = True
txtnktp.Enabled = True
cmbkk.Enabled = True
txtnke.Enabled = True
cmbj.Enabled = True
rdob.Enabled = True
rdoe.Enabled = True
txtht.Enabled = True
txtjb.Enabled = True
txtth.Enabled = True
End Sub

Sub tidakaktif()
txtnko.Enabled = False
txtnp.Enabled = False
txtnktp.Enabled = False
cmbkk.Enabled = False
txtnke.Enabled = False
cmbj.Enabled = False
rdob.Enabled = False
rdoe.Enabled = False
txtht.Enabled = False
txtjb.Enabled = False
txtth.Enabled = False
End Sub

Sub bersih()
txtnko.Text = ""
txtnp.Text = ""
txtnktp.Text = ""
cmbkk.Text = ""
txtnke.Text = ""
cmbj.Text = ""
txtht.Text = ""
txtjb.Text = ""
txtth.Text = ""
txtnke.SetFocus
End Sub


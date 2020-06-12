VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form formmasteradm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Menu Utama"
   ClientHeight    =   5085
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   120
      Top             =   4680
   End
   Begin ComctlLib.StatusBar stbar2 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   4590
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   3175
            MinWidth        =   3175
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdmember 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Daftar Member"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      Picture         =   "formmasteradm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmddatapemesan 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Data Pemesan"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      Picture         =   "formmasteradm.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdkereta 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Data Kereta"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      Picture         =   "formmasteradm.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   9480
      Top             =   4200
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FF00&
      Height          =   135
      Left            =   0
      TabIndex        =   9
      Top             =   4440
      Width           =   10095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MENU"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "WhatsApp (+62)83811274758 | E-Mail myticket@gmail.co.id"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   1320
      Width           =   6015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Jl. Raya Sukahati, No. 05 Cibinong Bogor Selatan"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   6855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Height          =   135
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   10095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Selamat Datang Admin | Mari Kita Utamakan Pelanggan   |"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   1680
      Width           =   6855
   End
   Begin VB.Menu laporan 
      Caption         =   "Laporan/Data"
   End
   Begin VB.Menu user 
      Caption         =   "User"
      Begin VB.Menu gpassword 
         Caption         =   "Ganti Paswword"
      End
   End
   Begin VB.Menu windows 
      Caption         =   "Windows"
      Begin VB.Menu about 
         Caption         =   "About Program"
      End
   End
   Begin VB.Menu keluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "formmasteradm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
formaboutprogram.Show
End Sub

Private Sub cmddatapemesan_Click()
frmpemesan.Show
frmpemesan.txtcari.SetFocus
Me.Hide
End Sub

Private Sub cmdkereta_Click()
frmkereta.Show
Me.Hide
End Sub

Private Sub cmdmember_Click()
Me.Hide
frmdatapemesan.Show
End Sub

Private Sub gpassword_Click()
formgpass.Show
Me.Visible = False
End Sub

Private Sub keluar_Click()
b = MsgBox("Apakah Anda Ingin Keluar ?", vbQuestion + vbOKCancel, "Informasi")
If b = vbOK Then
MsgBox "Sukses Anda Telah Keluar Dari Akun Anda", vbInformation, "Informasi"
Unload Me
formlogin.Show
End If
End Sub

Private Sub laporan_Click()
formlaporan.Show
formlaporan.cmdkembali.SetFocus
Me.Hide
End Sub

Private Sub timer1_timer()
Dim lng_chr As Integer
Dim r_chr, l_chr As String
lng_chr = Len(Label1.Caption)
l_chr = Left(Label1.Caption, 1)
r_chr = Right(Label1.Caption, lng_chr - 1)
Label1.Caption = r_chr + l_chr
End Sub

Private Sub Timer2_Timer()
formmasteradm.stbar2.Panels(3) = Time
End Sub

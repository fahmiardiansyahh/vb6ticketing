VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form menuutamapesan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELAMAT DATANG DI myTICKET !!!"
   ClientHeight    =   6555
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   14625
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmpesan.frx":0000
   ScaleHeight     =   6555
   ScaleWidth      =   14625
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdgpass 
      Caption         =   "&Ganti Password"
      Height          =   975
      Left            =   6960
      Picture         =   "frmpesan.frx":28EDF
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   0
      Top             =   2400
   End
   Begin VB.CommandButton cmdpesan 
      Caption         =   "&Pesan"
      Height          =   975
      Left            =   3600
      Picture         =   "frmpesan.frx":298E1
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton keluar 
      Caption         =   "&Keluar"
      Height          =   975
      Left            =   10440
      Picture         =   "frmpesan.frx":2A2E3
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   840
   End
   Begin ComctlLib.StatusBar stbar1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   5940
      Width           =   14625
      _ExtentX        =   25797
      _ExtentY        =   1085
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
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
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
      Left            =   6600
      TabIndex        =   9
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Line Line1 
      DrawMode        =   9  'Not Mask Pen
      X1              =   0
      X2              =   15360
      Y1              =   600
      Y2              =   600
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
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   2400
      Width           =   6495
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
      Left            =   4200
      TabIndex        =   5
      Top             =   1560
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   5280
      Width           =   15135
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Height          =   135
      Left            =   0
      TabIndex        =   3
      Top             =   2880
      Width           =   15255
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   15360
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lselamat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SELAMAT DATANG DI PROGRAM PEMESANAN TIKET KERETA API OLEH MyTIKET SILAHKAN KLIK PESAN UNTUK MEMESAN TIKET"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   15135
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15135
   End
End
Attribute VB_Name = "menuutamapesan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdgpass_Click()
gpasuser.Show
Me.Hide
End Sub

Private Sub cmdpesan_Click()
pesan.Show
menuutamapesan.Hide
End Sub

Private Sub keluar_Click()
X = MsgBox("Apakah Anda Ingin Keluar ?", vbQuestion + vbOKCancel, "Informasi")
If X = vbOK Then
MsgBox "Sukses Anda Telah Keluar Dari Akun Anda", vbInformation, "Informasi"
Unload Me
formlogin.Show
End If
End Sub

Private Sub timer1_timer()
lselamat.ForeColor = RGB(Rnd * 250, Rnd * 250, Rnd * 250)
If (lselamat.Left + lselamat.Width) <= 0 Then
lselamat.Left = Me.Width
End If
lselamat.Left = lselamat.Left - 100
End Sub

Private Sub Timer2_Timer()
menuutamapesan.stbar1.Panels(3) = Time
End Sub

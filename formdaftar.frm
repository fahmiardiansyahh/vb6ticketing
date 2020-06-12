VERSION 5.00
Begin VB.Form formdaftar 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daftar Gratis!!"
   ClientHeight    =   7245
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9435
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "formdaftar.frx":0000
   ScaleHeight     =   7245
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmddaftar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ya! Daftar Sekarang"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      Picture         =   "formdaftar.frx":28EDF
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6240
      Width           =   3855
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7920
      Top             =   7560
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lihat Kata Sandi"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6480
         TabIndex        =   14
         Top             =   5520
         Width           =   1695
      End
      Begin VB.TextBox txtkon 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   13
         Top             =   4920
         Width           =   4335
      End
      Begin VB.TextBox txtid 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   12
         Top             =   1920
         Width           =   4335
      End
      Begin VB.TextBox txtks 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   11
         Top             =   4080
         Width           =   4335
      End
      Begin VB.TextBox txtnb 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   10
         Top             =   3360
         Width           =   4335
      End
      Begin VB.TextBox txtnd 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   9
         Top             =   2640
         Width           =   4335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Daftar Sekarang! Mulai Perjalanan  Wisata dan Liburan Anda"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   7575
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Konfirmasi Kata Sandi"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   5040
         Width           =   3015
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kata Sandi"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Id Pemesan"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nama Belakang"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nama Depan"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Sosa"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Buat Akun myTICKET Anda Hari Ini!"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   600
         TabIndex        =   1
         Top             =   240
         Width           =   6975
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   8400
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Height          =   855
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Width           =   8415
      End
   End
   Begin VB.Label lmasuk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MASUK Sekarang!"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7200
      TabIndex        =   18
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label lsudah 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Sudah Punya Akun?"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   17
      Top             =   6600
      Width           =   1935
   End
End
Attribute VB_Name = "formdaftar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmddaftar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.cmddaftar.BackColor = vbGreen
End Sub

Private Sub Check1_Click()
If Check1 = 1 Then
txtks.PasswordChar = ""
txtkon.PasswordChar = ""
Else
txtks.PasswordChar = "*"
txtkon.PasswordChar = "*"
End If
End Sub
Sub bersih()
txtnd.Text = ""
txtnb.Text = ""
txtks.Text = ""
txtkon.Text = ""
Check1 = 0
End Sub

Private Sub cmddaftar_Click()
Call koneksi
If txtnd.Text = "" Then
MsgBox "Data Masih Ada yang Kosong Harap Di Isi", vbCritical, "Informasi"
ElseIf txtnb.Text = "" Then
MsgBox "Data Masih Ada yang Kosong Harap Di Isi", vbCritical, "Informasi"
ElseIf txtks.Text = "" Then
MsgBox "Data Masih Ada yang Kosong Harap Di Isi", vbCritical, "Informasi"
ElseIf txtkon.Text = "" Then
MsgBox "Data Masih Ada yang Kosong Harap Di Isi", vbCritical, "Informasi"
ElseIf txtkon.Text <> txtks.Text Then
MsgBox "Konfirmasi Kata Sandi Berbeda", vbOKOnly, "Informasi"
txtkon.SetFocus
Else
simpan = "Insert into tb_pemesan values ('" & txtid.Text & "','" & txtnd.Text + txtnb.Text & "','" & txtks.Text & "')"
KON.Execute simpan
MsgBox "Sukses Anda Telah Terdaftar Silahkan Login Untuk Memesan", vbOKOnly, "Informasi"
Call bersih
formlogin.Show
Unload Me
End If
End Sub
Private Sub Form_Load()
txtks.PasswordChar = "*"
txtkon.PasswordChar = "*"
Call koneksi
Call id
Call aktif
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.cmddaftar.BackColor = vbWhite
Me.lmasuk.ForeColor = vbRed
End Sub
Private Sub lmasuk_Click()
b = MsgBox("Apakah Anda Igin Membatalkan Pendaftaran ,Jika Ya Anda Akan Di Alihkan Ke Form Login", vbQuestion + vbOKCancel, "Informasi")
If b = vbOK Then
Unload Me
formlogin.Show
End If
End Sub

Private Sub lmasuk_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.lmasuk.ForeColor = vbBlue
End Sub

Private Sub timer1_timer()
Label9.ForeColor = RGB(Rnd * 250, Rnd * 250, Rnd * 250)
End Sub
Sub id()
koneksi
rspesan.Open ("select * from tb_pemesan where id_pemesan in(select max(id_pemesan)from tb_pemesan)order by id_pemesan Desc"), KON, adOpenKeyset

Dim id_pemesan As String * 5
Dim Hitung As Long
If rspesan.EOF Then
nopemesan = "PN" + "001"
txtid.Text = nopemesan
Else
Hitung = Right(rspesan!id_pemesan, 3) + 1
nopemesan = "PN" + Right("000" & Hitung, 3)
End If
txtid.Text = nopemesan
End Sub
Sub aktif()
txtid.Enabled = False
txtnd.Enabled = True
txtnb.Enabled = True
txtks.Enabled = True
txtkon.Enabled = True
End Sub
Private Sub txtkon_KeyPress(KeyAscii As Integer)
txtkon.MaxLength = 8
If KeyAscii = 13 Then
If Check1 = 1 Then
txtkon.PasswordChar = ""
Else
txtkon.PasswordChar = "*"
If txtks = "" Then
MsgBox "Password Konfirmasi Harus Di Isi ", vbCritical, "Informasi"
txtkon.SetFocus
Else
If txtkon.Text <> txtks.Text Then
MsgBox "Konfirmasi Kata Sandi Berbeda", vbOKOnly, "Informasi"
Else
cmddaftar.SetFocus
End If
End If
End If
End If
End Sub
Private Sub txtks_KeyPress(KeyAscii As Integer)
txtks.MaxLength = 8
If KeyAscii = 13 Then
If Check1 = 1 Then
txtks.PasswordChar = ""
Else
txtks.PasswordChar = "*"
If txtks = "" Then
MsgBox "Password Tidak Boleh Kosong Harap di Isi", vbCritical, "Informasi"
txtks.SetFocus
Else
txtkon.SetFocus
End If
End If
End If
End Sub
Private Sub txtnama_Change()
nama = txtnd.Text + txtnb.Text
txtnama.Text = nama
txtnama.Refresh
End Sub

Private Sub txtnb_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
If txtnb = "" Then
MsgBox "Nama Belakang Tidak Boleh Kosong ", vbCritical, "Informasi"
txtnb.SetFocus
Else
txtks.SetFocus
End If
End If
End Sub

Private Sub txtnd_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txtnd.MaxLength = 10
txtnb.MaxLength = 20
If KeyAscii = 13 Then
If txtnd.Text = "" Then
MsgBox "Nama Depan Tidak Boleh Kosong", vbCritical, "Informasi"
txtnd.SetFocus
Else
txtnb.SetFocus
End If
End If
End Sub

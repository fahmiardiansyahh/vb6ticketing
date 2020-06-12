VERSION 5.00
Begin VB.Form formlogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selamat Datang"
   ClientHeight    =   5355
   ClientLeft      =   105
   ClientTop       =   750
   ClientWidth     =   8280
   FillColor       =   &H000000C0&
   Icon            =   "formlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "formlogin.frx":0A02
   ScaleHeight     =   5355
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdmasuk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Masuk"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      Picture         =   "formlogin.frx":298E1
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Left            =   7320
      Top             =   5040
   End
   Begin VB.CheckBox clihat 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lihat Password"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5640
      TabIndex        =   6
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtpass 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4320
      TabIndex        =   5
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox txtuser 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4320
      TabIndex        =   4
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton cmdkeluar 
      BackColor       =   &H00008080&
      BeginProperty Font 
         Name            =   "Sosa"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Picture         =   "formlogin.frx":2A2E3
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   840
      Width           =   8415
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   5040
      Width           =   8415
   End
   Begin VB.Label ldaftar 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "DAFTAR Sekarang!"
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
      Left            =   6360
      TabIndex        =   8
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Belum Punya Akun?"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label lpas 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   2160
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lusername 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ID "
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
      Left            =   2160
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1080
      Left            =   600
      Picture         =   "formlogin.frx":2ACE5
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Label lsign 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SIGN IN"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   8295
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   840
      Width           =   8415
   End
   Begin VB.Menu login 
      Caption         =   "&Login"
   End
   Begin VB.Menu cetak 
      Caption         =   "&Cetak-Bukti-Pemesanan"
   End
End
Attribute VB_Name = "formlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b As Byte

Private Sub cetak_Click()
Unload Me
frmcek.Show
End Sub

Private Sub clihat_Click()
If clihat = 1 Then
txtpass.PasswordChar = ""
Else
txtpass.PasswordChar = "*"
End If
End Sub

Private Sub cmdkeluar_Click()
x = MsgBox("Apakah Anda Ingin Keluar ?", vbQuestion + vbOKCancel, "Informasi")
If x = vbOK Then
End
End If
End Sub
Private Sub cmdmasuk_Click()
Call koneksi
rspesan.Open "select * from tb_pemesan where id_pemesan ='" & txtuser.Text & "' and password='" & txtpass.Text & "'", KON
rsadmin.Open "select * from tb_admin where id_petugas ='" & txtuser.Text & "' and password='" & txtpass.Text & "'", KON
If rspesan.EOF And rsadmin.EOF Then
b = b + 1
If 1 - b = 0 Then
MsgBox "Kesempatan ke " & b & " salah" & Chr(13) & "password '" & txtpass.Text & "' tidak dikenal"
txtpas = ""
txtpass.SetFocus
ElseIf 2 - b = 0 Then
MsgBox "Kesempatan ke " & b & " salah" & Chr(13) & "password '" & txtpass.Text & "' tidak dikenal"
txtpas = ""
txtpass.SetFocus
ElseIf 3 - b = 0 Then
MsgBox "Kesempatan Ke " & b & " salah" & Chr(13) & "password '" & txtpass.Text & "' Tidak Di kenal" & Chr(13) & "Kesempatan Habis ,ulangi Dari awal"
Unload Me
End If
Else
If Not rspesan.EOF Then
frmSplash.Show
menuutamapesan.stbar1.Panels(1) = rspesan!id_pemesan
menuutamapesan.stbar1.Panels(2) = rspesan!nama_pemesan
Unload Me
End If
If Not rsadmin.EOF Then
MsgBox "Anda Login Sebagai Admin", vbInformation, "Informasi"
formmasteradm.Show
formmasteradm.stbar2.Panels(1) = rsadmin!id_petugas
formmasteradm.stbar2.Panels(2) = rsadmin!nama_petugas
Unload Me
End If
End If
End Sub

Private Sub Form_Load()
txtpass.PasswordChar = "*"
txtpass.MaxLength = 8
txtuser.MaxLength = 5
Label1.Caption = "          SELAMAT DATANG | SILAHKAN LOGIN UNTUK MELAKUKAN PEMESANAN"
Label8.Caption = "          DAFTAR SEKARANG!! Apabila Belum Mempunyai Akun!!"
Label8.Top = Label1.Top
Label1.AutoSize = True
Label8.AutoSize = True
Label8.FontSize = Label1.FontSize
Label1.Left = 0
Label8.Left = Me.Width
Timer1.Interval = 100
End Sub

Private Sub ldaftar_Click()
a = MsgBox("Apakah Anda Ingin Mendaftar ?", vbQuestion + vbOKCancel, "Informasi")
If a = vbOK Then
formdaftar.Show
formlogin.Hide
Else
formlogin.Show
End If
End Sub

Private Sub timer1_timer()
If Label1.Left > -Label1.Width Then
Label1.Left = Label1.Left - 50
Else
If Label8.Left < 0 Then
Label1.Left = Me.Width
End If
End If

If Label8.Left > -Label8.Width Then
Label8.Left = Label8.Left - 50
Else
If Label1.Left < 0 Then
Label8.Left = Me.Width
End If
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.ldaftar.ForeColor = vbRed
Me.cmdmasuk.BackColor = vbWhite
Me.cmdkeluar.BackColor = &HC0C0&
End Sub

Private Sub ldaftar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.ldaftar.ForeColor = vbBlue
End Sub

Private Sub cmdmasuk_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdmasuk.BackColor = vbGreen
End Sub

Private Sub cmdkeluar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdkeluar.BackColor = vbYellow
End Sub
Private Sub txtuser_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
Call koneksi
rspesan.Open "select id_pemesan from tb_pemesan where id_pemesan='" & txtuser & "'", KON
rsadmin.Open "select id_petugas from tb_admin where id_petugas='" & txtuser & "'", KON
If rspesan.EOF And rsadmin.EOF Then
a = a + 1
If 1 - a = 0 Then
MsgBox "Kesempatan Ke" & a & "salah" & Chr(13) & "nama'" & txtuser & "' Tidak Di Kenal "
txtuser = ""
txtuser.SetFocus
ElseIf 2 - a = 0 Then
MsgBox "Kesempatan ke " & a & "Salah" & Chr(13) & "nama'" & txtuser & "' Tidak Di Kenal "
txtuser = ""
txtuser.SetFocus
ElseIf 3 - a = 0 Then
MsgBox " Kesempatan ke " & a & "Salah" & Chr(13) & "nama'" & txtuser & "' Tidak Dikenal" & Chr(13) & "Kesempatan habis,ulangi dari awal"
Unload Me
End If
Else
txtuser.Enabled = False
txtpass.Enabled = True
txtpass.SetFocus
End If
End If
End Sub

VERSION 5.00
Begin VB.Form frmdatapemesan 
   BorderStyle     =   0  'None
   Caption         =   "Form Member"
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6165
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "&Hapus"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtidpemesan 
      Appearance      =   0  'Flat
      Height          =   450
      Left            =   2280
      TabIndex        =   12
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cmdkeluar 
      Caption         =   "&Keluar"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdubah 
      Caption         =   "&Ubah"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton caricmd 
      Caption         =   "&Cari"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.ComboBox caricmb 
      Height          =   450
      Left            =   2280
      TabIndex        =   7
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CheckBox clihat 
      Caption         =   "Lihat Kata Sandi"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   6
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox txtkatasandi 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2280
      TabIndex        =   5
      Top             =   3720
      Width           =   3255
   End
   Begin VB.TextBox txtnamapemesan 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2280
      TabIndex        =   4
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kata Sandi"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   600
      TabIndex        =   3
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Pemesan"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   2
      Top             =   3120
      Width           =   1380
   End
   Begin VB.Label lidpemesan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Id Pemesan"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   480
      TabIndex        =   1
      Top             =   2520
      Width           =   1050
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000000&
      BorderWidth     =   2
      Height          =   3975
      Left            =   120
      Top             =   1320
      Width           =   5895
   End
   Begin VB.Label lpelanggan 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FORM DATA MEMBER"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1260
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Height          =   5535
      Left            =   -120
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label lcari 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Member"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   13
      Top             =   1680
      Width           =   1695
   End
End
Attribute VB_Name = "frmdatapemesan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub caridata()
Call koneksi
rsmember.Open "Select * from tb_pemesan", KON
caricmb.Clear
caricmb.Refresh
Do While Not rsmember.EOF
caricmb.AddItem rsmember!id_pemesan
rsmember.MoveNext
Loop
End Sub
Private Sub caricmb_Click()
caricmd.Enabled = True
End Sub

Private Sub caricmd_Click()
Call koneksi
rsmember.Open "Select * from tb_pemesan where id_pemesan='" & Me.caricmb.Text & "'", KON
If Not rsmember.EOF Then
txtidpemesan = rsmember!id_pemesan
txtnamapemesan = rsmember!nama_pemesan
txtkatasandi = rsmember!Password
cmdubah.Enabled = True
cmdbatal.Enabled = True
cmdhapus.Enabled = True
ElseIf rsmember.EOF Then
MsgBox "Maaf Kode Tidak Di temukan", vbInformation, "Informasi"
End If
End Sub

Private Sub clihat_Click()
If clihat = 1 Then
txtkatasandi.PasswordChar = ""
Else
txtkatasandi.PasswordChar = "*"
End If
End Sub

Private Sub cmdbatal_Click()
Call bersih
End Sub

Private Sub cmdhapus_Click()
x = MsgBox("Yakin akan dihapus", vbYesNo)
If x = vbYes Then
Dim SQLHapus As String
SQLHapus = "Delete From tb_pemesan where id_pemesan= '" & Me.caricmb.Text & "'"
KON.Execute SQLHapus
Unload Me
frmdatapemesan.Show
Call nonaktif
End If
End Sub

Private Sub cmdkeluar_Click()
a = MsgBox("Apakah Anda Ingin Keluar ? Jika Ya Maka Anda Akan Kembali Ke Form Master Admin !", vbQuestion + vbOKCancel, "Informasi")
If a = vbOK Then
Unload Me
formmasteradm.Show
End If
End Sub

Private Sub cmdsimpan_Click()
If txtnamapemesan.Text = "" Or txtkatasandi.Text = "" Then
MsgBox "Maaf Data Tidak Dapat Di Simpan !!! Harap Isi Field Yang Kosong !!!", vbCritical, "Informasi"
Else
Dim sqlubah As String
sqlubah = "Update tb_pemesan Set nama_pemesan= '" & Me.txtnamapemesan & "', password='" & Me.txtkatasandi & "' where id_pemesan= '" & Me.caricmb.Text & "'"
KON.Execute sqlubah
MsgBox "Data Telah Di Update", vbInformation, "Informasi"
Unload Me
frmdatapemesan.Show
End If
End Sub

Private Sub cmdubah_Click()
txtnamapemesan.Enabled = True
txtkatasandi.Enabled = True
clihat.Enabled = True
cmdhapus.Enabled = False
cmdubah.Enabled = False
cmdkeluar.Enabled = False
caricmb.Enabled = False
caricmd.Enabled = False
cmdsimpan.Enabled = True
txtnamapemesan.SetFocus
End Sub

Private Sub Form_Load()
Call koneksi
Call caridata
Call nonaktif
txtnamapemesan.MaxLength = 30
txtkatasandi.MaxLength = 8
txtkatasandi.PasswordChar = "*"
End Sub
Sub nonaktif()
txtidpemesan.Enabled = False
txtnamapemesan.Enabled = False
txtkatasandi.Enabled = False
cmdsimpan.Enabled = False
cmdubah.Enabled = False
cmdbatal.Enabled = False
caricmd.Enabled = False
clihat.Enabled = False
cmdhapus.Enabled = False
End Sub
Sub aktif()
txtnamapemesan.Enabled = True
txtkatasandi.Enabled = True
End Sub
Sub bersih()
txtidpemesan.Text = ""
txtnamapemesan.Text = ""
txtkatasandi.Text = ""
caricmb.Text = ""
caricmd.Enabled = False
cmdhapus.Enabled = False
cmdubah.Enabled = False
cmdbatal.Enabled = False
cmdsimpan.Enabled = False
clihat.Enabled = False
caricmb.Enabled = True
cmdkeluar.Enabled = True
caricmb.SetFocus
End Sub
Private Sub txtnamapemesan_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
If txtnamapemesan.Text = "" Then
MsgBox "Harap Isi Nama Pemesan !!!", vbCritical, "Informasi"
txtnamapemesan.SetFocus
End If
End If
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmkereta 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11955
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dttiba 
      Height          =   375
      Left            =   8760
      TabIndex        =   37
      Top             =   4200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   103022593
      CurrentDate     =   42853
   End
   Begin VB.TextBox txtid 
      Height          =   375
      Left            =   3240
      TabIndex        =   36
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtharga 
      Height          =   375
      Left            =   8760
      TabIndex        =   33
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox txtjamtiba 
      Height          =   375
      Left            =   8760
      TabIndex        =   32
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox txtsttujuan 
      Height          =   375
      Left            =   8760
      TabIndex        =   31
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtkursi 
      Height          =   375
      Left            =   3240
      TabIndex        =   26
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox txtstberangkat 
      Height          =   375
      Left            =   3240
      TabIndex        =   22
      Top             =   6000
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtberangkat 
      Height          =   375
      Left            =   8760
      TabIndex        =   21
      Top             =   2400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   103153665
      CurrentDate     =   42715
   End
   Begin VB.TextBox txttujuan 
      Height          =   375
      Left            =   3240
      TabIndex        =   20
      Top             =   4200
      Width           =   1695
   End
   Begin VB.ComboBox cmbkelas 
      Height          =   390
      Left            =   3240
      TabIndex        =   16
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdkeluar 
      Caption         =   "&EXIT"
      Height          =   495
      Left            =   9240
      TabIndex        =   15
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   5640
      TabIndex        =   14
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdubah 
      Caption         =   "&Ubah"
      Height          =   495
      Left            =   2280
      TabIndex        =   13
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   7440
      TabIndex        =   12
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   3960
      TabIndex        =   11
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdtambah 
      Caption         =   "&Tambah"
      Height          =   495
      Left            =   480
      TabIndex        =   10
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "PENCARIAN BERDASARKAN ID KERETA"
      Height          =   975
      Left            =   2640
      TabIndex        =   8
      Top             =   1200
      Width           =   6495
      Begin VB.ComboBox cmbcari 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   4455
      End
      Begin VB.CommandButton cmdcari 
         Appearance      =   0  'Flat
         Caption         =   "&CARI"
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
         Left            =   4920
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.TextBox txtjamberangkat 
      Height          =   375
      Left            =   8760
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtpemberangkatan 
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtnama 
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ID KERETA"
      Height          =   375
      Left            =   1200
      TabIndex        =   35
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HARGA TIKET"
      Height          =   375
      Left            =   6720
      TabIndex        =   30
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "JAM TIBA"
      Height          =   375
      Left            =   6720
      TabIndex        =   29
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL TIBA"
      Height          =   375
      Left            =   6480
      TabIndex        =   28
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "STASIUN TUJUAN"
      Height          =   375
      Left            =   6600
      TabIndex        =   27
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Image Image4 
      Height          =   975
      Left            =   -120
      Picture         =   "frmkereta.frx":0000
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   13815
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Kereta"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   25
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label11 
      Caption         =   "Di Gunakan untuk Mengedit Jadwal Kereta"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   24
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "FORM INI UNTUK PENAMBAHAN DAN PENGURANGAN JADWAL KERETA"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   23
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "JAM KEBERANGKATAN"
      Height          =   375
      Left            =   6240
      TabIndex        =   19
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "STASIUN KEBERANGKATAN"
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "KURSI TERSEDIA"
      Height          =   495
      Left            =   1080
      TabIndex        =   17
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   5760
      X2              =   5760
      Y1              =   2280
      Y2              =   6600
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL KEBERANGKATAN"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "KELAS"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TUJUAN"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PEMBERANGKATAN"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label lnkereta 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA KERETA"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   120
      Picture         =   "frmkereta.frx":238B
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1320
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   0
      Picture         =   "frmkereta.frx":DBB2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13815
   End
End
Attribute VB_Name = "frmkereta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isi As Boolean
Sub caridata()
    Call koneksi
    rskereta.Open "select * from tb_kereta", KON
    cmbcari.Clear
    cmbcari.Refresh
    Do While Not rskereta.EOF
    cmbcari.AddItem rskereta!id_kereta
    rskereta.MoveNext
    Loop
    End Sub
Private Sub KosongkanText()
    Me.txtnama = ""
    Me.txtid = ""
    Me.txtpemberangkatan = ""
    Me.txttujuan = ""
    Me.cmbkelas = ""
    Me.txtkursi = ""
    Me.txtstberangkat = ""
    Me.txtjamberangkat = ""
    Me.txtsttujuan = ""
    Me.txtjamtiba = ""
    Me.txtharga = ""
End Sub
Private Sub SiapIsi()
    Me.txtnama.Enabled = True
    Me.txtpemberangkatan.Enabled = True
    Me.txttujuan.Enabled = True
    Me.txtid.Enabled = True
    Me.cmbkelas.Enabled = True
    Me.txttujuan.Enabled = True
    Me.cmbkelas.Enabled = True
    Me.txtkursi.Enabled = True
    Me.txtstberangkat.Enabled = True
    Me.dtberangkat.Enabled = True
    Me.txtjamberangkat.Enabled = True
    Me.txtsttujuan.Enabled = True
    Me.dttiba.Enabled = True
    Me.txtjamtiba.Enabled = True
    Me.txtharga.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Me.txtnama.Enabled = False
    Me.txtpemberangkatan.Enabled = False
    Me.txtid.Enabled = False
    Me.txttujuan.Enabled = False
    Me.cmbkelas.Enabled = False
    Me.txttujuan.Enabled = False
    Me.cmbkelas.Enabled = False
    Me.txtkursi.Enabled = False
    Me.txtstberangkat.Enabled = False
    Me.dtberangkat.Enabled = False
    Me.txtjamberangkat.Enabled = False
    Me.txtsttujuan.Enabled = False
    Me.dttiba.Enabled = False
    Me.txtjamtiba.Enabled = False
    Me.txtharga.Enabled = False
End Sub
 
Private Sub KondisiAwal()
    KosongkanText
    TidakSiapIsi
    Me.cmdbatal.Enabled = False
    Me.cmdsimpan.Enabled = False
    Me.cmdhapus.Enabled = False
    Me.cmdubah.Enabled = False
    cmdcari.Enabled = False
    cmdtambah.Enabled = True
    cmdkeluar.Enabled = True
    cmbcari = ""
End Sub

Private Sub cmbcari_Click()
cmdcari.Enabled = True
End Sub

Private Sub cmbkelas_Click()
txtkursi.SetFocus
End Sub

Private Sub cmdbatal_Click()
KondisiAwal
cmbcari.Enabled = True
End Sub
 
Private Sub cmdcari_Click()
Call koneksi
rskereta.Open "Select * from tb_kereta where id_kereta='" & Me.cmbcari.Text & "'", KON
If Not rskereta.EOF Then
txtid = rskereta!id_kereta
txtnama = rskereta!nama_kereta
txtpemberangkatan = rskereta!pemberangkatan
txttujuan = rskereta!tujuan
cmbkelas = rskereta!kelas
txtkursi = rskereta!kursi_tersedia
txtstberangkat = rskereta!stasiun_keberangkatan
dtberangkat = rskereta!tanggal_keberangkatan
txtjamberangkat = rskereta!jam_keberangkatan
txtsttujuan = rskereta!stasiun_tujuan
dttiba.Value = rskereta!tanggal_tiba
txtjamtiba = rskereta!jam_tiba
txtharga = rskereta!harga
cmdubah.Enabled = True
cmdhapus.Enabled = True
cmdbatal.Enabled = True
cmdtambah.Enabled = False
cmdkeluar.Enabled = False
ElseIf rskereta.EOF Then
MsgBox "Kode Tidak Ada"
KondisiAwal
End If
End Sub

Private Sub cmdhapus_Click()
x = MsgBox("Yakin akan dihapus", vbYesNo)
If x = vbYes Then
Dim SQLHapus As String
SQLHapus = "Delete From tb_kereta where id_kereta= '" & Me.cmbcari.Text & "'"
KON.Execute SQLHapus
Unload Me
frmkereta.Show
KondisiAwal
End If
End Sub

Private Sub cmdkeluar_Click()
k = MsgBox("Apakah Anda Ingin Keluar Dari Form Ini?Jika Ya anda Akan Di alihkan Ke Form Master", vbQuestion + vbOKCancel, "Informasi")
If k = vbOK Then
Unload Me
formmasteradm.Show
End If
End Sub
Private Sub cmdsimpan_Click()
Call koneksi
If txtid.Text = "" Then
MsgBox "Data Masih Ada Yang Kosong Harap Di Isi", vbCritical, "Informasi"
ElseIf txtnama.Text = "" Then
MsgBox "Data Masih Ada Yang Kosong Harap Di Isi", vbCritical, "Informasi"
ElseIf txtpemberangkatan.Text = "" Then
MsgBox "Data Masih Ada Yang Kosong Harap Di Isi", vbCritical, "Informasi"
ElseIf txttujuan.Text = "" Then
MsgBox "Data Masih Ada Yang Kosong Harap Di Isi", vbCritical, "Informasi"
ElseIf cmbkelas.Text = "" Then
MsgBox "Data Masih Ada Yang Kosong Harap Di Isi", vbCritical, "Informasi"
ElseIf txtkursi.Text = "" Then
MsgBox "Data Masih Ada Yang Kosong Harap Di Isi", vbCritical, "Informasi"
ElseIf txtstberangkat.Text = "" Then
MsgBox "Data Masih Ada Yang Kosong Harap Di Isi", vbCritical, "Informasi"
ElseIf txtjamberangkat.Text = "" Then
MsgBox "Data Masih Ada Yang Kosong Harap Di Isi", vbCritical, "Informasi"
ElseIf txtsttujuan.Text = "" Then
MsgBox "Data Masih Ada Yang Kosong Harap Di Isi", vbCritical, "Informasi"
ElseIf txtjamtiba.Text = "" Then
MsgBox "Data Masih Ada Yang Kosong Harap Di Isi", vbCritical, "Informasi"
ElseIf txtharga.Text = "" Then
MsgBox "Data Masih Ada Yang Kosong Harap Di Isi", vbCritical, "Informasi"
Else
If isi = True Then
Dim simpan As String
SQLSimpan = "Insert Into tb_kereta values('" & Me.txtid & "','" & Me.txtnama & "','" & Me.txtpemberangkatan & "','" & Me.txttujuan & "','" & Me.cmbkelas.Text & "','" & Me.txtkursi & "','" & Me.txtstberangkat & "','" & Format(Me.dtberangkat.Value, "YYYY/MM/DD") & "','" & Me.txtjamberangkat & "','" & Me.txtsttujuan & "','" & Format(Me.dttiba.Value, "YYYY/MM/DD") & "','" & Me.txtjamtiba & "','" & Me.txtharga & "')"
KON.Execute SQLSimpan
MsgBox "Data Berhasil Di Simpan", vbInformation, "Informasi"
Unload Me
frmkereta.Show
ElseIf isi = False Then
Dim SQLEdit As String
SQLEdit = "Update tb_kereta Set kursi_tersedia= '" & Me.txtkursi & "', tanggal_keberangkatan='" & Format(Me.dtberangkat.Value, "YYYY/MM/DD") & "',jam_keberangkatan='" & Me.txtjamberangkat & "',tanggal_tiba='" & Format(Me.dttiba.Value, "YYYY/MM/DD") & "',jam_tiba='" & Me.txtjamtiba & "',harga='" & Me.txtharga & "' where id_kereta= '" & Me.txtid & "'"
KON.Execute SQLEdit
MsgBox "Data Berhasil Di Simpan", vbInformation, "Informasi"
Unload Me
frmkereta.Show
End If
KondisiAwal
cmbcari.Enabled = True
End If
End Sub

Private Sub cmdtambah_Click()
isi = True
SiapIsi
txtid.SetFocus
cmbcari.Enabled = False
cmdtambah.Enabled = False
cmdkeluar.Enabled = False
cmdcari.Enabled = False
Me.cmdbatal.Enabled = True
Me.cmdsimpan.Enabled = True
End Sub

 
Private Sub cmdubah_Click()
isi = False
SiapIsi
txtid.SetFocus
txtid.Enabled = True
txtnama.Enabled = True
Me.txtpemberangkatan.Enabled = True
Me.txttujuan.Enabled = True
Me.txtstberangkat.Enabled = True
Me.txtsttujuan.Enabled = True
Me.cmbkelas.Enabled = True
cmbcari.Enabled = False
cmdcari.Enabled = False
cmdhapus.Enabled = False
cmdubah.Enabled = False
Me.txtkursi.Enabled = True
Me.dtberangkat.Enabled = True
Me.txtjamberangkat.Enabled = True
Me.dttiba.Enabled = True
Me.txtjamtiba.Enabled = True
Me.txtharga.Enabled = True
cmdsimpan.Enabled = True
cmdbatal.Enabled = True
End Sub
Private Sub Form_Load()
Call koneksi
Call caridata
Call batas
cmbkelas.AddItem "EKONOMI AC"
cmbkelas.AddItem "BISNIS"
cmbkelas.AddItem "EKSEKUTIF"
KondisiAwal
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdcari.Enabled = True
End If
End Sub
Sub batas()
Me.txtid.MaxLength = 9
Me.txtnama.MaxLength = 20
Me.txtpemberangkatan.MaxLength = 15
Me.txttujuan.MaxLength = 15
Me.txtkursi.MaxLength = 3
Me.txtstberangkat.MaxLength = 20
Me.txtjamberangkat.MaxLength = 5
Me.txtsttujuan.MaxLength = 20
Me.txtjamtiba.MaxLength = 5
Me.txtharga.MaxLength = 10
End Sub
Private Sub txtharga_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
cmdsimpan.SetFocus
End If
End Sub

Private Sub txtid_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
txtnama.SetFocus
End If
End Sub
Private Sub txtjamberangkat_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
txtsttujuan.SetFocus
End If
End Sub
Private Sub txtjamtiba_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
txtharga.SetFocus
End If
End Sub

Private Sub txtkursi_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
txtstberangkat.SetFocus
End If
End Sub

Private Sub txtnama_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
txtpemberangkatan.SetFocus
End If
End Sub
Private Sub txtpemberangkatan_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
txttujuan.SetFocus
End If
End Sub
Private Sub txtstberangkat_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
dtberangkat.SetFocus
End If
End Sub
Private Sub txtsttujuan_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
dttiba.SetFocus
End If
End Sub
Private Sub txttujuan_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
cmbkelas.SetFocus
End If
End Sub

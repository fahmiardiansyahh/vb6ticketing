VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form formprint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Priview"
   ClientHeight    =   10140
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   17010
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport cr 
      Left            =   14400
      Top             =   9480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox kembali 
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14160
      TabIndex        =   40
      Top             =   7080
      Width           =   2175
   End
   Begin VB.TextBox uangbayar 
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14160
      TabIndex        =   38
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   2040
      Top             =   9600
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15000
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   9360
      Width           =   1335
   End
   Begin VB.CommandButton cmdbatal 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   9480
      Width           =   1335
   End
   Begin VB.TextBox sttujuan 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   34
      Top             =   8760
      Width           =   3015
   End
   Begin VB.TextBox stberangkat 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   33
      Top             =   8760
      Width           =   3375
   End
   Begin VB.TextBox total 
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14160
      TabIndex        =   32
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox harga 
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14160
      TabIndex        =   31
      Top             =   4920
      Width           =   2175
   End
   Begin VB.TextBox jumlahpesan 
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14160
      TabIndex        =   30
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox tanggaltiba 
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   29
      Top             =   6000
      Width           =   2655
   End
   Begin VB.TextBox jamtiba 
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   28
      Top             =   6720
      Width           =   2655
   End
   Begin VB.TextBox jamberangkat 
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   27
      Top             =   5280
      Width           =   2655
   End
   Begin VB.TextBox tanggalberangkat 
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   26
      Top             =   4560
      Width           =   2655
   End
   Begin VB.TextBox tanggalpemesanan 
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12840
      TabIndex        =   25
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox idpemesan 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   24
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox kelas 
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   23
      Top             =   6480
      Width           =   2655
   End
   Begin VB.TextBox namakereta 
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   22
      Top             =   5640
      Width           =   2655
   End
   Begin VB.TextBox idkereta 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   21
      Top             =   4920
      Width           =   2655
   End
   Begin VB.TextBox namapemesan 
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   20
      Top             =   3240
      Width           =   5055
   End
   Begin VB.TextBox notransaksi 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   19
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   11640
      X2              =   11640
      Y1              =   3960
      Y2              =   7680
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5880
      X2              =   5880
      Y1              =   3960
      Y2              =   7680
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   600
      X2              =   16560
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   480
      X2              =   16440
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Kembali"
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   39
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Uang Bayar"
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   37
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "LIHAT APAKAH ADA YANG SALAH?"
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   18
      Top             =   2040
      Width           =   5415
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH PESAN"
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   17
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "HARGA"
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   16
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   15
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "TGL. TIBA"
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "JAM. BERANGKAT"
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   13
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NO. TRANSAKSI"
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ST. TUJUAN"
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   11
      Top             =   8160
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "TGL. BERANGKAT"
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "JAM. TIBA"
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ST. KEBERANGKATAN"
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   8160
      Width           =   2895
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "KELAS"
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA KERETA"
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ID KERETA"
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA PEMESAN"
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ID PEMESAN"
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL PEMESANAN"
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12600
      TabIndex        =   2
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   2280
      Left            =   480
      Picture         =   "formprint.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16005
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008000&
      Height          =   10335
      Left            =   16440
      TabIndex        =   1
      Top             =   -240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Height          =   10215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "formprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub panggil()
notransaksi.Text = formbayar.txtnotransaksi.Text
namapemesan.Text = menuutamapesan.stbar1.Panels(2).Text
idpemesan.Text = menuutamapesan.stbar1.Panels(1).Text
tanggalpemesanan.Text = transaksi.txttanggal.Caption
idkereta.Text = transaksi.kodekereta.Caption
namakereta.Text = transaksi.lnamakereta.Caption
kelas.Text = transaksi.lkelas.Caption
tanggalberangkat.Text = transaksi.tanggalberangkat.Caption
jamberangkat.Text = pesan.grid.Columns(8)
jamtiba.Text = pesan.grid.Columns(11)
tanggaltiba.Text = transaksi.tanggalsampai.Caption
jumlahpesan.Text = transaksi.jumlahtiket.Caption
harga.Text = transaksi.hargatiket.Caption
total.Text = transaksi.total.Caption
uangbayar.Text = formbayar.TxtJumlahBayar.Text
kembali.Text = formbayar.TxtKembali.Text
stberangkat.Text = transaksi.berangkat.Caption
sttujuan.Text = transaksi.tujuan.Caption
End Sub


Private Sub cmdbatal_Click()
If cmdbatal.Caption = "&SELESAI" Then
deletetemp = "Delete From tb_temp where notransaksi= '" & Me.notransaksi.Text & "'"
KON.Execute deletetemp
MsgBox "Transaksi Anda Telah Berhasil !!! Anda Akan Di Bawa Ke Menu utama ", vbInformation, "Informasi"
Call Form_Unload(1)
menuutamapesan.Show
ElseIf cmdbatal.Caption = "&BATAL" Then
x = MsgBox("Apakah Anda Ingin Membatalkan Seluruh Proses Transaksi?", vbQuestion + vbOKCancel, "Informasi")
If x = vbOK Then
hapus = "Delete From tb_transaksi where notransaksi= '" & Me.notransaksi.Text & "'"
KON.Execute hapus
hapusdetail = "Delete From tb_detailtransaksi where notransaksi= '" & Me.notransaksi.Text & "'"
KON.Execute hapusdetail
hapustemp = "Delete From tb_temp where notransaksi= '" & Me.notransaksi.Text & "'"
KON.Execute hapustemp
tambah = "update tb_kereta set kursi_tersedia=kursi_tersedia +'" & Val(transaksi.jumlahtiket.Caption) & "' where id_kereta = '" & transaksi.kodekereta.Caption & " '"
KON.Execute (tambah)
MsgBox "Transaksi Sukses Di Batalkan !!!!", vbInformation, "Informasi"
Call Form_Unload(1)
menuutamapesan.Show
End If
End If
End Sub

Private Sub cmdprint_Click()
If cmdprint.Caption = "&PRINT" Then
Call cetak
cmdbatal.Caption = "&SELESAI"
End If
End Sub

Private Sub Form_Load()
cmdbatal.Caption = "&BATAL"
cmdprint.Caption = "&PRINT"
Call koneksi
Call panggil
Call nonaktif
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload formbayar
Unload transaksi
Unload Me
Unload pesan
End Sub


Private Sub timer1_timer()
Label20.ForeColor = RGB(Rnd * 250, Rnd * 120, Rnd * 350)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.cmdbatal.BackColor = vbWhite
Me.cmdprint.BackColor = vbWhite
End Sub

Private Sub cmdbatal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.cmdbatal.BackColor = vbRed
End Sub

Private Sub cmdprint_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdprint.BackColor = vbGreen
End Sub
Sub nonaktif()
notransaksi.Enabled = False
namapemesan.Enabled = False
idpemesan.Enabled = False
tanggalpemesanan.Enabled = False
idkereta.Enabled = False
namakereta.Enabled = False
kelas.Enabled = False
tanggalberangkat.Enabled = False
jamberangkat.Enabled = False
jamtiba.Enabled = False
tanggaltiba.Enabled = False
jumlahpesan.Enabled = False
harga.Enabled = False
total.Enabled = False
uangbayar.Enabled = False
kembali.Enabled = False
stberangkat.Enabled = False
sttujuan.Enabled = False
End Sub
Sub cetak()
Call koneksi
cr.ReportFileName = App.Path & "\rtransaksi.rpt"
cr.WindowState = crptNormal
cr.RetrieveDataFiles
cr.Action = 1
End Sub

VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmcek 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cetak Bukti Pemesanan"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14145
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
   ScaleHeight     =   8055
   ScaleWidth      =   14145
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport cr1 
      Left            =   120
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdlaporan 
      Caption         =   "&CETAK"
      Height          =   495
      Left            =   3600
      TabIndex        =   35
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton cmdkeluar 
      Appearance      =   0  'Flat
      Caption         =   "&KELUAR"
      Height          =   495
      Left            =   12120
      TabIndex        =   34
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdbatal 
      Appearance      =   0  'Flat
      Caption         =   "&BATAL"
      Height          =   495
      Left            =   7080
      TabIndex        =   33
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox txtkembalian 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   11760
      TabIndex        =   32
      Top             =   6360
      Width           =   2295
   End
   Begin VB.TextBox txtuangbayar 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   11760
      TabIndex        =   31
      Top             =   5280
      Width           =   2295
   End
   Begin VB.TextBox txttotalharga 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   11880
      TabIndex        =   30
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox txtjumlahbeli 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   12600
      TabIndex        =   29
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtharga 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   11760
      TabIndex        =   28
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtsttujuan 
      Height          =   375
      Left            =   7440
      TabIndex        =   27
      Top             =   6360
      Width           =   2415
   End
   Begin VB.TextBox txtstberangkat 
      Height          =   375
      Left            =   7440
      TabIndex        =   26
      Top             =   5280
      Width           =   2415
   End
   Begin VB.TextBox txtkelas 
      Height          =   375
      Left            =   7440
      TabIndex        =   25
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox txtnamakereta 
      Height          =   375
      Left            =   7440
      TabIndex        =   24
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox txtidkereta 
      Height          =   375
      Left            =   7440
      TabIndex        =   23
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox txtnamapemesan 
      Height          =   375
      Left            =   2640
      TabIndex        =   22
      Top             =   5640
      Width           =   2415
   End
   Begin VB.TextBox txttanggalpesan 
      Height          =   375
      Left            =   2640
      TabIndex        =   21
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox txtidpemesan 
      Height          =   375
      Left            =   2640
      TabIndex        =   20
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox txtnotransaksi 
      Height          =   375
      Left            =   2640
      TabIndex        =   19
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "PENCARIAN BERDASARKAN KODE"
      Height          =   975
      Left            =   2880
      TabIndex        =   2
      Top             =   1200
      Width           =   6495
      Begin VB.CommandButton cmdcari 
         Appearance      =   0  'Flat
         Caption         =   "&CARI"
         Height          =   495
         Left            =   4800
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtcari 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "TR-99999999"
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "frmcek.frx":0000
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   14175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "KEMBALIAN"
      Height          =   495
      Left            =   10200
      TabIndex        =   18
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UANG BAYAR"
      Height          =   495
      Left            =   10080
      TabIndex        =   17
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label ltotal 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL HARGA"
      Height          =   495
      Left            =   10200
      TabIndex        =   16
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label ljumlahbeli 
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH BELI"
      Height          =   495
      Left            =   10320
      TabIndex        =   15
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lharga 
      BackStyle       =   0  'Transparent
      Caption         =   "HARGA"
      Height          =   375
      Left            =   10560
      TabIndex        =   14
      Top             =   2760
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   9960
      X2              =   9960
      Y1              =   2280
      Y2              =   6960
   End
   Begin VB.Label ltujuan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "STASIUN TUJUAN"
      Height          =   255
      Left            =   5520
      TabIndex        =   13
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label lstasiunberangkat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "STASIUN KEBERANGKATAN"
      Height          =   735
      Left            =   5520
      TabIndex        =   12
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label lkelas 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "KELAS"
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label lnamakereta 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA KERETA"
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label lidkereta 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ID KERETA"
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   5280
      X2              =   5280
      Y1              =   2280
      Y2              =   6960
   End
   Begin VB.Label lnamapemesan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA PEMESAN"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label tgljual 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL PESAN"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label lidpemesan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ID PEMESAN"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lidtiket 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No TRANSAKSI"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lcarikode 
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Id Transaksi :"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Image lsimbol 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   120
      Picture         =   "frmcek.frx":238B
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1680
   End
   Begin VB.Label lfcetak 
      BackStyle       =   0  'Transparent
      Caption         =   "CETAK BUKTI PEMESANAN"
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
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   0
      Picture         =   "frmcek.frx":DBB2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14295
   End
End
Attribute VB_Name = "frmcek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isi As Boolean
Private Sub KosongkanText()
    Me.txtidpemesan = ""
    Me.txtnotransaksi = ""
    Me.txttanggalpesan = ""
    Me.txtnamapemesan = ""
    Me.txtidkereta = ""
    Me.txtnamakereta = ""
    Me.txtkelas = ""
    Me.txtstberangkat = ""
    Me.txtsttujuan = ""
    Me.txtharga = ""
    Me.txtjumlahbeli = ""
    Me.txttotalharga = ""
    txtuangbayar = ""
    txtkembalian = ""
End Sub
Private Sub SiapIsi()
    Me.txtidpemesan.Enabled = True
    Me.txtnotransaksi.Enabled = True
    Me.txttanggalpesan.Enabled = True
    Me.txtnamapemesan.Enabled = True
    Me.txtidkereta.Enabled = True
    Me.txtnamakereta.Enabled = True
    Me.txtkelas.Enabled = True
    Me.txtstberangkat.Enabled = True
    Me.txtsttujuan.Enabled = True
    Me.txtharga.Enabled = True
    Me.txtjumlahbeli.Enabled = True
    Me.txttotalharga.Enabled = True
    Me.txtuangbayar.Enabled = True
    Me.txtkembalian.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Me.txtidpemesan.Enabled = False
    Me.txtnotransaksi.Enabled = False
    Me.txttanggalpesan.Enabled = False
    Me.txtnamapemesan.Enabled = False
    Me.txtidkereta.Enabled = False
    Me.txtnamakereta.Enabled = False
    Me.txtkelas.Enabled = False
    Me.txtstberangkat.Enabled = False
    Me.txtsttujuan.Enabled = False
    Me.txtharga.Enabled = False
    Me.txtjumlahbeli.Enabled = False
    Me.txttotalharga.Enabled = False
    Me.txtuangbayar.Enabled = False
    Me.txtkembalian.Enabled = False
End Sub
 
Private Sub KondisiAwal()
    KosongkanText
    TidakSiapIsi
    Me.cmdbatal.Enabled = False
    cmdcari.Enabled = False
    cmdkeluar.Enabled = True
    cmdlaporan.Enabled = False
    txtcari.Text = ""
End Sub
Private Sub cmdbatal_Click()
KondisiAwal
cmdcari.Enabled = False
txtcari.Enabled = True
txtcari.SetFocus
End Sub
 
Private Sub cmdcari_Click()
Call koneksi
rstransaksi.Open "select * from tb_transaksi where notransaksi= '" & txtcari.Text & "'", KON
If Not rstransaksi.EOF Then
txtcari.Enabled = False
cmdcari.Enabled = False
cmdbatal.Enabled = True
cmdlaporan.Enabled = True
cmdlaporan.SetFocus
txtidpemesan = rstransaksi!id_pemesan
txtnotransaksi = rstransaksi!notransaksi
txttanggalpesan = rstransaksi!tgl_jual
txtnamapemesan = rstransaksi!nama_pemesan
txtidkereta = rstransaksi!id_kereta
txtnamakereta = rstransaksi!nama_kereta
txtkelas = rstransaksi!kelas
txtstberangkat = rstransaksi!pemberangkatan
txtsttujuan = rstransaksi!tujuan
txtharga = rstransaksi!harga
txtjumlahbeli = rstransaksi!jumlah_pesanan
txttotalharga = rstransaksi!sub_total
txtuangbayar = rstransaksi!uang_bayar
txtkembalian = rstransaksi!kembalian
ElseIf rstransaksi.EOF Then
MsgBox "Kode Tidak Ada"
KondisiAwal
End If
End Sub
Private Sub cmdkeluar_Click()
Unload Me
formlogin.Show
End Sub

Private Sub cmdlaporan_Click()
Call koneksi
cr1.SelectionFormula = "{tb_transaksi.notransaksi}='" & txtnotransaksi.Text & "'"
cr1.ReportFileName = App.Path & "\lrtransaksi.rpt"
cr1.WindowState = crptNormal
cr1.RetrieveDataFiles
cr1.Action = 1
End Sub

Private Sub Form_Load()
Call koneksi
KondisiAwal
End Sub
Private Sub txtcari_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
cmdcari.Enabled = True
End If
End Sub



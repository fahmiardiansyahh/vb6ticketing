VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmpemesan 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13245
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
   ScaleHeight     =   7890
   ScaleWidth      =   13245
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid 
      Height          =   4575
      Left            =   0
      TabIndex        =   4
      Top             =   2280
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   8070
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdkeluar 
      Caption         =   "&KELUAR"
      Height          =   495
      Left            =   11400
      TabIndex        =   1
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "PENCARIAN BERDASARKAN KODE"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   12975
      Begin VB.OptionButton Option4 
         Caption         =   "Kelas"
         Height          =   375
         Left            =   5760
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         Caption         =   "ID Kereta"
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "ID Pemesan"
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "No.Transaksi"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2055
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
         Left            =   7800
         TabIndex        =   3
         ToolTipText     =   "TR-99999999"
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   -240
      Picture         =   "frmpemesan.frx":0000
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   13575
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "FORM DATA PEMESAN"
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
      Left            =   4800
      TabIndex        =   2
      Top             =   0
      Width           =   3495
   End
   Begin VB.Image lsimbol 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   120
      Picture         =   "frmpemesan.frx":238B
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1920
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   -120
      Picture         =   "frmpemesan.frx":DBB2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13575
   End
End
Attribute VB_Name = "frmpemesan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub aktifgrid()
MSFlexGrid.Cols = 15
MSFlexGrid.RowHeightMin = 300
MSFlexGrid.Col = 0
MSFlexGrid.Row = 0
MSFlexGrid.Text = "NO"
MSFlexGrid.CellFontBold = True
MSFlexGrid.ColWidth(0) = 500
MSFlexGrid.AllowUserResizing = flexResizeColumns
MSFlexGrid.CellAlignment = flexAlignCenterCenter
MSFlexGrid.Col = 1
MSFlexGrid.Row = 0
MSFlexGrid.Text = "No. Transaksi"
MSFlexGrid.CellFontBold = True
MSFlexGrid.ColWidth(1) = 2000
MSFlexGrid.AllowUserResizing = flexResizeColumns
MSFlexGrid.CellAlignment = flexAlignCenterCenter
MSFlexGrid.Col = 2
MSFlexGrid.Row = 0
MSFlexGrid.Text = "Tgl Memesan"
MSFlexGrid.CellFontBold = True
MSFlexGrid.ColWidth(2) = 2000
MSFlexGrid.AllowUserResizing = flexResizeColumns
MSFlexGrid.CellAlignment = flexAlignCenterCenter
MSFlexGrid.Col = 3
MSFlexGrid.Row = 0
MSFlexGrid.Text = "ID Pemesan"
MSFlexGrid.CellFontBold = True
MSFlexGrid.ColWidth(3) = 2000
MSFlexGrid.AllowUserResizing = flexResizeColumns
MSFlexGrid.CellAlignment = flexAlignCenterCenter
MSFlexGrid.Col = 4
MSFlexGrid.Row = 0
MSFlexGrid.Text = "Nama Pemesan "
MSFlexGrid.CellFontBold = True
MSFlexGrid.ColWidth(4) = 2000
MSFlexGrid.AllowUserResizing = flexResizeColumns
MSFlexGrid.CellAlignment = flexAlignCenterCenter
MSFlexGrid.Col = 5
MSFlexGrid.Row = 0
MSFlexGrid.Text = "ID Kereta"
MSFlexGrid.CellFontBold = True
MSFlexGrid.ColWidth(5) = 2300
MSFlexGrid.AllowUserResizing = flexResizeColumns
MSFlexGrid.CellAlignment = flexAlignCenterCenter
MSFlexGrid.Col = 6
MSFlexGrid.Row = 0
MSFlexGrid.Text = "Nama Kereta"
MSFlexGrid.CellFontBold = True
MSFlexGrid.ColWidth(6) = 2500
MSFlexGrid.AllowUserResizing = flexResizeColumns
MSFlexGrid.CellAlignment = flexAlignCenterCenter
MSFlexGrid.Col = 7
MSFlexGrid.Row = 0
MSFlexGrid.Text = "Pemberangkatan"
MSFlexGrid.CellFontBold = True
MSFlexGrid.ColWidth(7) = 2500
MSFlexGrid.AllowUserResizing = flexResizeColumns
MSFlexGrid.CellAlignment = flexAlignCenterCenter
MSFlexGrid.Col = 8
MSFlexGrid.Row = 0
MSFlexGrid.Text = "Tujuan"
MSFlexGrid.CellFontBold = True
MSFlexGrid.ColWidth(8) = 2000
MSFlexGrid.AllowUserResizing = flexResizeColumns
MSFlexGrid.CellAlignment = flexAlignCenterCenter
MSFlexGrid.Col = 9
MSFlexGrid.Row = 0
MSFlexGrid.Text = "Kelas"
MSFlexGrid.CellFontBold = True
MSFlexGrid.ColWidth(9) = 1700
MSFlexGrid.AllowUserResizing = flexResizeColumns
MSFlexGrid.CellAlignment = flexAlignCenterCenter
MSFlexGrid.Col = 10
MSFlexGrid.Row = 0
MSFlexGrid.Text = "Harga"
MSFlexGrid.CellFontBold = True
MSFlexGrid.ColWidth(10) = 2000
MSFlexGrid.AllowUserResizing = flexResizeColumns
MSFlexGrid.CellAlignment = flexAlignCenterCenter
MSFlexGrid.Col = 11
MSFlexGrid.Row = 0
MSFlexGrid.Text = "Jumlah Beli"
MSFlexGrid.CellFontBold = True
MSFlexGrid.ColWidth(11) = 1700
MSFlexGrid.AllowUserResizing = flexResizeColumns
MSFlexGrid.CellAlignment = flexAlignCenterCenter
MSFlexGrid.Col = 12
MSFlexGrid.Row = 0
MSFlexGrid.Text = "Sub Total"
MSFlexGrid.CellFontBold = True
MSFlexGrid.ColWidth(12) = 2500
MSFlexGrid.AllowUserResizing = flexResizeColumns
MSFlexGrid.CellAlignment = flexAlignCenterCenter
MSFlexGrid.Col = 13
MSFlexGrid.Row = 0
MSFlexGrid.Text = "Uang Bayar"
MSFlexGrid.CellFontBold = True
MSFlexGrid.ColWidth(13) = 2500
MSFlexGrid.AllowUserResizing = flexResizeColumns
MSFlexGrid.CellAlignment = flexAlignCenterCenter
MSFlexGrid.Col = 14
MSFlexGrid.Row = 0
MSFlexGrid.Text = "Kembalian"
MSFlexGrid.CellFontBold = True
MSFlexGrid.ColWidth(14) = 2000
MSFlexGrid.AllowUserResizing = flexResizeColumns
MSFlexGrid.CellAlignment = flexAlignCenterCenter
End Sub

Private Sub cmdkeluar_Click()
Unload Me
formmasteradm.Show
End Sub

Private Sub Form_Load()
Call koneksi
Call TampilGrid
End Sub
Sub TampilGrid()
MSFlexGrid.Clear
Call aktifgrid
MSFlexGrid.Rows = 2
Baris = 0
Call koneksi
rstransaksi.Open "SELECT * FROM tb_transaksi ORDER BY notransaksi ASC", KON, adOpenDynamic, adLockOptimistic
If rstransaksi.BOF Then
Exit Sub
Else
With rstransaksi
.MoveFirst
Do While Not .EOF
Baris = Baris + 1
MSFlexGrid.Rows = Baris + 1
MSFlexGrid.TextMatrix(Baris, 0) = Baris
MSFlexGrid.TextMatrix(Baris, 1) = !notransaksi
MSFlexGrid.TextMatrix(Baris, 2) = !tgl_jual
MSFlexGrid.TextMatrix(Baris, 3) = !id_pemesan
MSFlexGrid.TextMatrix(Baris, 4) = !nama_pemesan
MSFlexGrid.TextMatrix(Baris, 5) = !id_kereta
MSFlexGrid.TextMatrix(Baris, 6) = !nama_kereta
MSFlexGrid.TextMatrix(Baris, 7) = !pemberangkatan
MSFlexGrid.TextMatrix(Baris, 8) = !tujuan
MSFlexGrid.TextMatrix(Baris, 9) = !kelas
MSFlexGrid.TextMatrix(Baris, 10) = !harga
MSFlexGrid.TextMatrix(Baris, 11) = !jumlah_pesanan
MSFlexGrid.TextMatrix(Baris, 12) = !sub_total
MSFlexGrid.TextMatrix(Baris, 13) = !uang_bayar
MSFlexGrid.TextMatrix(Baris, 14) = !kembalian
.MoveNext
Loop
End With
End If
rstransaksi.Close
End Sub
Private Sub txtcari_Change()
If Option1.Value = False And Option2.Value = False And Option3.Value = False And Option4.Value = False Then
MsgBox "Pilih Opsi Pencarian Terlebih dulu", vbExclamation, "Perhatian"
Exit Sub
End If
MSFlexGrid.Clear
Call aktifgrid
MSFlexGrid.Rows = 2
Baris = 0
Call koneksi
If Option1.Value = True Then
rstransaksi.Open "SELECT * FROM tb_transaksi WHERE notransaksi LIKE '%" & txtcari.Text & "%'", KON, adOpenDynamic, adLockOptimistic
ElseIf Option2.Value = True Then
rstransaksi.Open "SELECT * FROM tb_transaksi WHERE id_pemesan LIKE '%" & txtcari.Text & "%'", KON, adOpenDynamic, adLockOptimistic
ElseIf Option3.Value = True Then
rstransaksi.Open "SELECT * FROM tb_transaksi WHERE id_kereta LIKE '%" & txtcari.Text & "%'", KON, adOpenDynamic, adLockOptimistic
ElseIf Option4.Value = True Then
rstransaksi.Open "SELECT * FROM tb_transaksi WHERE kelas LIKE '%" & txtcari.Text & "%'", KON, adOpenDynamic, adLockOptimistic
End If
If rstransaksi.BOF Then
Exit Sub
Else
With rstransaksi
.MoveFirst
Do While Not .EOF
Baris = Baris + 1
MSFlexGrid.Rows = Baris + 1
MSFlexGrid.TextMatrix(Baris, 0) = Baris
MSFlexGrid.TextMatrix(Baris, 1) = !notransaksi
MSFlexGrid.TextMatrix(Baris, 2) = !tgl_jual
MSFlexGrid.TextMatrix(Baris, 3) = !id_pemesan
MSFlexGrid.TextMatrix(Baris, 4) = !nama_pemesan
MSFlexGrid.TextMatrix(Baris, 5) = !id_kereta
MSFlexGrid.TextMatrix(Baris, 6) = !nama_kereta
MSFlexGrid.TextMatrix(Baris, 7) = !pemberangkatan
MSFlexGrid.TextMatrix(Baris, 8) = !tujuan
MSFlexGrid.TextMatrix(Baris, 9) = !kelas
MSFlexGrid.TextMatrix(Baris, 10) = !harga
MSFlexGrid.TextMatrix(Baris, 11) = !jumlah_pesanan
MSFlexGrid.TextMatrix(Baris, 12) = !sub_total
MSFlexGrid.TextMatrix(Baris, 13) = !uang_bayar
MSFlexGrid.TextMatrix(Baris, 14) = !kembalian
.MoveNext
Loop
End With
End If
rstransaksi.Close
End Sub


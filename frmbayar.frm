VERSION 5.00
Begin VB.Form formbayar 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7275
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
   Moveable        =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtid 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   2040
      TabIndex        =   18
      Top             =   1800
      Width           =   4815
   End
   Begin VB.TextBox txtnotransaksi 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   2040
      TabIndex        =   17
      Top             =   1080
      Width           =   4815
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "Batal"
      Height          =   495
      Left            =   5280
      TabIndex        =   14
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "&Proses"
      Height          =   495
      Left            =   3480
      TabIndex        =   13
      Top             =   7320
      Width           =   1575
   End
   Begin VB.TextBox TxtKembali 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   2040
      TabIndex        =   12
      Top             =   6360
      Width           =   4815
   End
   Begin VB.TextBox TxtJumlahBayar 
      Appearance      =   0  'Flat
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
      Left            =   2040
      TabIndex        =   10
      Top             =   5640
      Width           =   4815
   End
   Begin VB.TextBox TxtGrandTotal 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   2040
      TabIndex        =   8
      Top             =   4920
      Width           =   4815
   End
   Begin VB.TextBox TxtPPN 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   2040
      TabIndex        =   6
      Text            =   "0%"
      Top             =   3960
      Width           =   4815
   End
   Begin VB.TextBox TxtDiskon 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   2040
      TabIndex        =   4
      Text            =   "0%"
      Top             =   3240
      Width           =   4815
   End
   Begin VB.TextBox TxtTotalPembelian 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   2040
      TabIndex        =   2
      Top             =   2520
      Width           =   4815
   End
   Begin VB.Label txtidpemesan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID Pemesan"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   360
      TabIndex        =   16
      Top             =   1920
      Width           =   1020
   End
   Begin VB.Label lno 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No.Transaksi"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   360
      TabIndex        =   15
      Top             =   1200
      Width           =   1155
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kembali"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   480
      TabIndex        =   11
      Top             =   6480
      Width           =   690
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Bayar"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   360
      TabIndex        =   9
      Top             =   5760
      Width           =   1155
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   360
      TabIndex        =   7
      Top             =   5040
      Width           =   1035
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   120
      X2              =   7080
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pajak (ppn)"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Diskon (%)"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   360
      TabIndex        =   3
      Top             =   3360
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Pembelian"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   1410
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSAKSI PEMBAYARAN"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "formbayar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbatal_Click()
n = MsgBox("ApakahAnda Ingin Membatalkan Pesanan Anda ?", vbQuestion + vbOKCancel, "Informasi")
If n = vbOK Then
Call Form_Unload(1)
pesan.Show
End If
End Sub
Private Sub cmdkeluar_Click()
If TxtJumlahBayar.Text = "" Then
MsgBox "Jumlah Bayar Harus Di Isi !!!", vbCritical, "Informasi"
Else
If Val(TxtJumlahBayar.Text) < Val(TxtGrandTotal.Text) Then
MsgBox "Jumlah bayar kurang", vbCritical, "Informasi"
TxtJumlahBayar.SetFocus
Else
sqlsimpannotapenjualan = "insert into tb_transaksi values('" & formbayar.txtnotransaksi.Text & "','" & Format(transaksi.txttanggal.Caption, "YYYY/MM/DD") & "','" & txtid.Text & "','" & menuutamapesan.stbar1.Panels(2).Text & "','" & transaksi.kodekereta.Caption & "','" & transaksi.lnamakereta.Caption & "','" & transaksi.lberangkat.Caption & "','" & transaksi.ltujuan.Caption & "','" & transaksi.lkelas.Caption & "','" & transaksi.hargatiket.Caption & "','" & transaksi.jumlahtiket.Caption & "','" & TxtGrandTotal.Text & "','" & TxtJumlahBayar.Text & "','" & TxtKembali.Text & "')"
KON.Execute sqlsimpannotapenjualan
simpandetail = "insert into tb_detailtransaksi values('" & formbayar.txtnotransaksi.Text & "','" & transaksi.jumlahtiket.Caption & "','" & transaksi.total.Caption & "','" & transaksi.kodekereta.Caption & "')"
KON.Execute simpandetail
masuktemp = "insert into tb_temp values('" & formbayar.txtnotransaksi.Text & "','" & Format(transaksi.txttanggal.Caption, "YYYY/MM/DD") & "','" & menuutamapesan.stbar1.Panels(2) & "','" & transaksi.kodekereta.Caption & "','" & transaksi.lnamakereta.Caption & "','" & transaksi.lkelas.Caption & "','" & transaksi.lberangkat.Caption & "','" & transaksi.ltujuan.Caption & "','" & transaksi.jumlahtiket.Caption & "','" & transaksi.hargatiket.Caption & "','" & TxtGrandTotal.Text & "','" & TxtJumlahBayar.Text & "','" & TxtKembali.Text & "')"
KON.Execute masuktemp
kurang = "update tb_kereta set kursi_tersedia=kursi_tersedia -'" & Val(transaksi.jumlahtiket.Caption) & "' where id_kereta = '" & transaksi.kodekereta.Caption & " '"
KON.Execute (kurang)
MsgBox "Transaksi Sukses", vbInformation, "INFORMASI"
transaksi.Hide
Me.Hide
formprint.Show
End If
End If
End Sub

Private Sub Form_Load()
With transaksi
TxtGrandTotal.Text = .total.Caption
TxtTotalPembelian.Text = .total.Caption
Call koneksi
Call idpemesan
Call notransaksi
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload transaksi
Unload Me
End Sub

Private Sub TxtJumlahBayar_Change()
GrandTotal = Val(TxtGrandTotal.Text)
uangbayar = Val(TxtJumlahBayar.Text)
TxtKembali.Text = uangbayar - GrandTotal
End Sub
Private Sub TxtJumlahBayar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If TxtJumlahBayar.Text = "" Then
MsgBox "Jumlah bayar harus diisi", vbCritical, "Informasi"
TxtJumlahBayar.SetFocus
Else
If Val(TxtJumlahBayar.Text) < Val(TxtGrandTotal.Text) Then
MsgBox "Jumlah bayar kurang", vbInformation, "Informasi"
TxtJumlahBayar.SetFocus
Else
cmdkeluar_Click
End If
End If
End If
End Sub
Sub idpemesan()
With menuutamapesan
txtid.Text = .stbar1.Panels(1).Text
End With
End Sub
Sub notransaksi()
Call koneksi
rstransaksi.Open "select * from tb_transaksi where notransaksi in(select max(notransaksi) from tb_transaksi)order by notransaksi asc", KON
rstransaksi.Requery
Dim urut As String
Dim Hitung As Long
With rstransaksi
If .EOF Then
urut = "TR-" + Format(Date, "YYMMDD") + "01"
txtnotransaksi = urut
Else
If Mid(!notransaksi, 4, 6) <> Format(Date, "YYMMDD") Then
urut = "TR-" + Format(Date, "YYMMDD") + "01"
Else
Hitung = Right(!notransaksi, 2) + 1
urut = "TR-" + Format(Date, "YYMMDD") + Right("00" & Hitung, 2)
End If
End If
txtnotransaksi.Text = urut
End With
End Sub

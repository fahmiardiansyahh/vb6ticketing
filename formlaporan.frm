VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form formlaporan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4770
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
   ScaleHeight     =   8040
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   720
   End
   Begin VB.CommandButton cmdkembali 
      Height          =   375
      Left            =   120
      Picture         =   "formlaporan.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000004&
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Text            =   "Tanggal Awal"
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Bulanan"
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      Top             =   6480
      Width           =   4575
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Text            =   "Tahun"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Text            =   "Bulan"
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cbulan 
         Height          =   450
         Left            =   2280
         TabIndex        =   14
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox ctahun 
         Height          =   450
         Left            =   2280
         TabIndex        =   13
         Top             =   840
         Width           =   2175
      End
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000004&
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Text            =   "Tanggal Akhir"
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Mingguan"
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   4575
      Begin Crystal.CrystalReport crpemesanan 
         Left            =   1800
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox cmingguanawal 
         Height          =   450
         Left            =   2280
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox cmingguanakhir 
         Height          =   450
         Left            =   2280
         TabIndex        =   9
         Top             =   840
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Harian"
      DragIcon        =   "formlaporan.frx":058A
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   4575
      Begin Crystal.CrystalReport crkereta 
         Left            =   1800
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Text            =   "Tanggal"
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox charian 
         Height          =   450
         Left            =   2280
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Laporan Bulanan"
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Laporan Mingguan"
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label lharian 
      BackStyle       =   0  'Transparent
      Caption         =   "Laporan Harian  "
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label klikkereta 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   """KLIK DISINI"""
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label ldatakereta 
      BackStyle       =   0  'Transparent
      Caption         =   "Laporan Data Kereta  :"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "LAPORAN"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "formlaporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbulan_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Call bersih
End If
End Sub

Private Sub charian_Click()
crpemesanan.SelectionFormula = "Totext ({tb_transaksi.tgl_jual})='" & charian & "'"
crpemesanan.ReportFileName = App.Path & "\rharian.rpt"
crpemesanan.WindowState = crptNormal
crpemesanan.RetrieveDataFiles
crpemesanan.Action = 1
End Sub

Private Sub charian_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Call bersih
End If
End Sub

Private Sub cmdkembali_Click()
Unload Me
formmasteradm.Show
End Sub

Private Sub cmingguanakhir_Click()
If cmingguanawal = "" Then
MsgBox "Tanggal Awal Kosong", , "Informasi"
cmingguanawal.SetFocus
Exit Sub
End If
crpemesanan.SelectionFormula = "{tb_transaksi.tgl_jual} in date (" & cmingguanawal.Text & ") to date (" & cmingguanakhir.Text & ")"
crpemesanan.ReportFileName = App.Path & "\rmingguan.rpt"
crpemesanan.WindowState = crptNormal
crpemesanan.RetrieveDataFiles
crpemesanan.Action = 1
End Sub

Private Sub cmingguanakhir_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Call bersih
End If
End Sub

Private Sub cmingguanawal_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Call bersih
End If
End Sub

Private Sub ctahun_Click()
Call koneksi
rstransaksi.Open "select * from tb_transaksi where month(tgl_jual)=' " & Val(cbulan) & " ' and year(tgl_jual)=' " & (ctahun) & " ' ", KON
If rstransaksi.EOF Then
MsgBox "Data Tidak Ditemukan", vbExclamation, "UPS!!"
Exit Sub
cbulan.SetFocus
End If
crpemesanan.SelectionFormula = "Month({tb_transaksi.tgl_jual})=" & Val(cbulan.Text) & " and Year ({tb_transaksi.tgl_jual})=" & Val(ctahun.Text)
crpemesanan.ReportFileName = App.Path & "\rbulanan.rpt"
crpemesanan.WindowState = crptNormal
crpemesanan.RetrieveDataFiles
crpemesanan.Action = 1
End Sub

Private Sub ctahun_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Call bersih
End If
End Sub

Private Sub Form_Load()
Call koneksi
rstransaksi.Open "select distinct tgl_jual from tb_transaksi order by 1", KON
rstransaksi.Requery
Do Until rstransaksi.EOF
charian.AddItem rstransaksi!tgl_jual
cmingguanawal.AddItem Format(rstransaksi!tgl_jual, "YYYY, MM, DD")
cmingguanakhir.AddItem Format(rstransaksi!tgl_jual, "YYYY, MM, DD")
rstransaksi.MoveNext
Loop
For i = 1 To 12
cbulan.AddItem i
Next i
For i = 10 To 20
ctahun.AddItem 2000 + i
Next i
End Sub
Private Sub klikkereta_Click()
crkereta.ReportFileName = App.Path & "\rkereta.rpt"
crkereta.WindowState = crptNormal
crkereta.RetrieveDataFiles
crkereta.Action = 1
End Sub

Private Sub timer1_timer()
klikkereta.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub
Sub bersih()
cmdkembali.SetFocus
charian.Text = ""
cmingguanawal.Text = ""
cmingguanakhir.Text = ""
cbulan.Text = ""
ctahun.Text = ""
End Sub

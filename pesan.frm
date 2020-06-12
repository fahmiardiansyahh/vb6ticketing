VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form pesan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SILAHKAN PILIH JENIS KERETA"
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14970
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "pesan.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   14970
   StartUpPosition =   2  'CenterScreen
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
      Left            =   0
      Picture         =   "pesan.frx":28EDF
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtjumlah 
      Height          =   480
      Left            =   12840
      TabIndex        =   12
      Top             =   3120
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   4575
      Left            =   0
      TabIndex        =   11
      Top             =   4320
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   8070
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   20
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbkelas 
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
      Left            =   8640
      TabIndex        =   10
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CommandButton cmdcari 
      Caption         =   "&Cari"
      Height          =   975
      Left            =   12960
      Picture         =   "pesan.frx":298E1
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   1335
   End
   Begin VB.ComboBox cmbstasiuntn 
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
      Left            =   4920
      TabIndex        =   3
      Top             =   2520
      Width           =   2415
   End
   Begin VB.ComboBox cmbstasiunkn 
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
      Left            =   1200
      TabIndex        =   2
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label ljadwal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "JADWAL KERETA YANG TELAH DI UPDATE"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   15135
   End
   Begin VB.Label lselamat1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SILAHKAN PILIH JENIS KERETA YANG INGIN DI PESAN"
      Height          =   855
      Left            =   1920
      TabIndex        =   6
      Top             =   480
      Width           =   11895
   End
   Begin VB.Label lkelas 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kelas"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      TabIndex        =   5
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label ltujuan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pilih Tujuan"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   1
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label lpilihsk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pilih Keberangkatan"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label llkelas 
      BackColor       =   &H0080FF80&
      Height          =   1215
      Left            =   -120
      TabIndex        =   4
      Top             =   1920
      Width           =   12615
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
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   3720
      Width           =   15255
   End
End
Attribute VB_Name = "pesan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub tampilkereta()
Call koneksi
rskereta.CursorLocation = adUseClient
rskereta.Open "select * from tb_kereta order by id_kereta", KON
If Not rskereta.EOF Then
With rskereta
With grid
.Refresh
End With
End With
End If
Set grid.DataSource = rskereta
End Sub
Private Sub cmdcari_Click()
Call koneksi
With grid
rskereta.CursorLocation = adUseClient
rskereta.Open "select * from tb_kereta where pemberangkatan='" & cmbstasiunkn.Text & "' and tujuan='" & cmbstasiuntn.Text & "' and kelas='" & cmbkelas.Text & "'", KON
If rskereta.EOF Then
MsgBox "Data Kereta Tidak Di Temukan", vbInformation, "Informasi"
Call tampilkereta
Call bersihcari
ElseIf Not rskereta.EOF Then
Call aktifgrid
With rskereta
With grid
Set .DataSource = rskereta
.Refresh
End With
End With
End If
End With
End Sub
Sub kelas()
cmbkelas.AddItem "EKONOMI AC"
cmbkelas.AddItem "BISNIS"
cmbkelas.AddItem "EKSEKUTIF"
End Sub

Private Sub cmdkeluar_Click()
Unload Me
menuutamapesan.Show
End Sub

Private Sub Form_Load()
Call tampilkereta
Call koneksi
Call stasiun
Call tujuan
Call kelas
txtjumlah.Visible = False
Call nonaktifgrid
End Sub
Sub stasiun()
Call koneksi
    rskereta.Open "select * from tb_kereta", KON
    cmbstasiunkn.Refresh
    Do While Not rskereta.EOF
    cmbstasiunkn.AddItem rskereta!pemberangkatan
    rskereta.MoveNext
    Loop
End Sub
Sub tujuan()
Call koneksi
rskereta.Open "Select * from tb_kereta", KON
cmbstasiuntn.Refresh
Do While Not rskereta.EOF
cmbstasiuntn.AddItem rskereta!tujuan
rskereta.MoveNext
Loop
End Sub
Sub bersihcari()
cmbstasiunkn.Text = ""
cmbstasiuntn.Text = ""
cmbkelas.Text = ""
cmbstasiunkn.SetFocus
End Sub
Private Sub grid_DblClick()
Call koneksi
Dim jumlah As String
With transaksi
jumlah = InputBox("Masukan Jumlah Tiket Yang Anda Ingin Pesan ?", "Input")
txtjumlah.Text = jumlah
If jumlah = "" Then
MsgBox "Anda Harus Memasukan Jumlah Tiket !!!", vbCritical, "Informasi"
Else
rskereta.Open "select * from tb_kereta where id_kereta='" & grid.Columns(0).Text & "'", KON
If Val(txtjumlah.Text) > rskereta!kursi_tersedia Then
MsgBox "Hanya tersedia  " + grid.Columns(5).Text + " Kursi Saja"
Exit Sub
ElseIf jumlah = 0 Then
MsgBox "Jumlah Kursi Telah Habis Terjual Atau Anda Salah Memasukan Jumlah Angka !!! Silahkan Pilih Kereta Jenis Lain", vbCritical, "Informasi"
Unload Me
pesan.Show
Else
.lnamakereta = grid.Columns(1).Text
.kodekereta = grid.Columns(0).Text
.lberangkat = grid.Columns(2).Text
.ltujuan = grid.Columns(3).Text
.hargatiket = grid.Columns(12).Text
.berangkat = grid.Columns(6).Text
.tujuan = grid.Columns(9).Text
.tanggalberangkat = grid.Columns(7).Text
.tanggalsampai = grid.Columns(10).Text
.lkelas = grid.Columns(4).Text
Unload Me
transaksi.Show
End If
End If
End With
End Sub
Sub aktifgrid()
grid.Enabled = True
End Sub
Sub nonaktifgrid()
grid.Enabled = False
End Sub

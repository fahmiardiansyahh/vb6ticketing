VERSION 5.00
Begin VB.Form transaksi 
   BackColor       =   &H8000000C&
   Caption         =   "TRANSAKSI"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16950
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
   ScaleHeight     =   7410
   ScaleWidth      =   16950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdbatal 
      Caption         =   "&BATAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14040
      TabIndex        =   12
      Top             =   6000
      Width           =   2415
   End
   Begin VB.CommandButton cmdbayar 
      Caption         =   "&PROSES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11160
      TabIndex        =   11
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   17000
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   240
         Top             =   5160
      End
      Begin VB.Label ltujuan 
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
         Left            =   10200
         TabIndex        =   23
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lberangkat 
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
         Left            =   8040
         TabIndex        =   22
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label txttanggal 
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
         Left            =   14040
         TabIndex        =   21
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label tanggal 
         Caption         =   "TANGGAL"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12600
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lkelas 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   19
         Top             =   2040
         Width           =   4575
      End
      Begin VB.Label tanggalsampai 
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
         Height          =   375
         Left            =   3120
         TabIndex        =   17
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Label tanggalberangkat 
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
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label berangkat 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   10
         Top             =   3600
         Width           =   4215
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Sosa"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label tujuan 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   8
         Top             =   2880
         Width           =   4215
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "o"
         BeginProperty Font 
            Name            =   "Sosa"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "Sosa"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2880
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label total 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   9720
         TabIndex        =   4
         Top             =   1440
         Width           =   10335
      End
      Begin VB.Label hargatiket 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   7080
         TabIndex        =   3
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label jumlahtiket 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4680
         TabIndex        =   2
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label lnamakereta 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   1
         Top             =   1440
         Width           =   4575
      End
      Begin VB.Line Line1 
         X1              =   -2400
         X2              =   13440
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Image Image1 
         Height          =   1215
         Left            =   120
         Picture         =   "frmtransaksi.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2760
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "PEMBAYARAN TIKET"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   3960
         TabIndex        =   6
         Top             =   120
         Width           =   4095
      End
      Begin VB.Label kodekereta 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8040
         TabIndex        =   13
         Top             =   120
         Width           =   4095
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   18
      Top             =   4560
      Width           =   4215
   End
   Begin VB.Label ltanggal 
      BackColor       =   &H8000000A&
      Height          =   495
      Left            =   8160
      TabIndex        =   15
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   8640
      TabIndex        =   14
      Top             =   720
      Width           =   4095
   End
End
Attribute VB_Name = "transaksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, y As String

Private Sub cmdbatal_Click()
c = MsgBox("Apakah Anda Ingin Membatalkan Pesanan Anda ?", vbQuestion + vbOKCancel, "Informasi")
If c = vbOK Then
Unload Me
pesan.Show
End If
End Sub

Private Sub cmdbayar_Click()
c = MsgBox("Apakah Anda Yakin Mau Memesan Jika Ya Anda Tidak Dapat Kembali Ke form Sebelumnya?", vbQuestion + vbOKCancel, "Informasi")
If c = vbOK Then
MsgBox "Anda Harus Bayar Terlebih Dahulu", vbInformation + vbOKOnly, "Informasi"
formbayar.Show
Me.Enabled = False
cmdbatal.Enabled = False
cmdbayar.Enabled = False
End If
End Sub

Private Sub Form_Load()
With pesan
hargatiket.Caption = .txtjumlah.Text
jumlahtiket.Caption = .txtjumlah.Text
Call koneksi
x = Val(.grid.Columns(12).Text)
y = Val(hargatiket.Caption)
total.Caption = x * y
End With
End Sub


Private Sub timer1_timer()
txttanggal.Caption = Format(Date, "YYYY/MM/DD")
End Sub

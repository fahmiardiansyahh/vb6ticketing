VERSION 5.00
Begin VB.Form formgpass 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Ganti Kata Sandi"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtuser 
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
      Left            =   3360
      TabIndex        =   13
      Top             =   840
      Width           =   3855
   End
   Begin VB.CheckBox chkpass 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lihat Kata Sandi"
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
      Left            =   5520
      TabIndex        =   12
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtpwlama 
      Alignment       =   2  'Center
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
      Left            =   3360
      TabIndex        =   11
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox txtpwbaru 
      Alignment       =   2  'Center
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
      Left            =   3360
      TabIndex        =   10
      Top             =   2640
      Width           =   3855
   End
   Begin VB.CheckBox chkpassbaru 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lihat Kata Sandi baru"
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
      Left            =   5040
      TabIndex        =   9
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox txtkon 
      Alignment       =   2  'Center
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
      Left            =   3360
      TabIndex        =   8
      Top             =   3360
      Width           =   3855
   End
   Begin VB.CommandButton cmdkeluar 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      Picture         =   "formgpass.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton cmdganti 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&GANTI"
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
      Left            =   2280
      Picture         =   "formgpass.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Sosa"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "GANTI KATA SANDI"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Konfirmasi Kata Sandi"
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
      Left            =   240
      TabIndex        =   5
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Kata Sandi Baru"
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
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Kata Sandi"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
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
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   5880
      Width           =   10455
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
Attribute VB_Name = "formgpass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkpass_Click()
If chkpass = 1 Then
txtpwlama.PasswordChar = ""
Else
txtpwlama.PasswordChar = "*"
End If
End Sub

Private Sub chkpassbaru_Click()
If chkpassbaru = 1 Then
txtpwbaru.PasswordChar = ""
txtkon.PasswordChar = ""
Else
txtpwbaru.PasswordChar = "*"
txtkon.PasswordChar = "*"
End If
End Sub

Private Sub cmdganti_Click()
Call koneksi
If txtkon.Text <> txtpwbaru.Text Then
MsgBox "Konfirmasi Password Berbeda"
txtkon.SetFocus
Else
rsadmin.Open "Update tb_admin set password='" & txtkon.Text & "' where id_petugas='" & txtuser.Text & "'", KON
MsgBox "Password Telah Di Update"
Call Form_Activate
txtpwlama = ""
txtkon = ""
txtpwbaru = ""
End If
End Sub

Private Sub cmdkeluar_Click()
Unload Me
formmasteradm.Show
End Sub
Private Sub Form_Activate()
For Each k In Me.Controls
If TypeOf k Is TextBox Then
k.Enabled = False
End If
Next
txtuser = formmasteradm.stbar2.Panels(1).Text
txtpwlama.Enabled = True
txtpwlama.SetFocus
txtpwlama.PasswordChar = "*"
txtpwlama.MaxLength = 8
txtpwbaru.MaxLength = 8
txtkon.MaxLength = 8
txtpwbaru.PasswordChar = "*"
txtkon.PasswordChar = "*"
End Sub
Private Sub txtpwbaru_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtkon.Enabled = True
txtkon.SetFocus
End If
End Sub

Private Sub txtpwlama_KeyPress(KeyAscii As Integer)
Call koneksi
If KeyAscii = 13 Then
rsadmin.Open "Select * from tb_admin where password='" & txtpwlama.Text & "'", KON
If rsadmin.EOF Then
MsgBox "password" + txtpwlama.Text + "Tidak ada"
Else
txtpwbaru.Enabled = True
txtpwbaru.SetFocus
End If
End If
End Sub

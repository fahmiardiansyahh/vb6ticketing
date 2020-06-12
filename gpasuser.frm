VERSION 5.00
Begin VB.Form gpasuser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ganti Password"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5970
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
   ScaleHeight     =   5610
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdkeluar 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      Picture         =   "gpasuser.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5160
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
      Left            =   1680
      Picture         =   "gpasuser.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4440
      Width           =   2895
   End
   Begin VB.CheckBox chkpassbaru 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lihat  Sandi baru"
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
      Left            =   4200
      TabIndex        =   12
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox txtkon 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   3480
      Width           =   3855
   End
   Begin VB.TextBox txtpwbaru 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   2640
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
      Left            =   4200
      TabIndex        =   7
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtpwlama 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1560
      Width           =   3855
   End
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
      Left            =   2040
      TabIndex        =   4
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Konfirmasi Sandi"
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
      Left            =   0
      TabIndex        =   10
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Kata Sandi Baru"
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
      Left            =   0
      TabIndex        =   8
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
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
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
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1455
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
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   3855
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
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Height          =   615
      Left            =   -600
      TabIndex        =   0
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "gpasuser"
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
rspesan.Open "Update tb_pemesan set password='" & txtkon.Text & "' where id_pemesan='" & txtuser.Text & "'", KON
MsgBox "Password Telah Di Update"
Call Form_Activate
txtpwlama = ""
txtkon = ""
txtpwbaru = ""
End If
End Sub

Private Sub cmdkeluar_Click()
Unload Me
menuutamapesan.Show
End Sub
Private Sub Form_Activate()
For Each k In Me.Controls
If TypeOf k Is TextBox Then
k.Enabled = False
End If
Next
txtuser = menuutamapesan.stbar1.Panels(1).Text
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
rspesan.Open "Select * from tb_pemesan where password='" & txtpwlama.Text & "'", KON
If rspesan.EOF Then
MsgBox "password" + txtpwlama.Text + "Tidak ada"
Else
txtpwbaru.Enabled = True
txtpwbaru.SetFocus
End If
End If
End Sub


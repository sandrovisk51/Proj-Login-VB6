VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   8460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnEntrar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Entrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      MaskColor       =   &H00C0FFFF&
      TabIndex        =   2
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtEmail 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   7335
   End
   Begin VB.TextBox txtSenha 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   600
      TabIndex        =   0
      Top             =   1920
      Width           =   7335
   End
   Begin VB.Line Line2 
      X1              =   600
      X2              =   8040
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   8040
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblSenha 
      BackStyle       =   0  'Transparent
      Caption         =   "SENHA :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "EMAIL :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   8175
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnEntrar_Click()

If txtEmail.Text = "" Then
    lblMsg.ForeColor = &HFF&
    lblMsg.Caption = "Informe a Email."
    txtEmail.SetFocus
    Exit Sub
End If

If txtSenha.Text = "" Then
    lblMsg.ForeColor = &HFF&
    lblMsg.Caption = "Informe a Senha."
    txtSenha.SetFocus
    Exit Sub
End If
    
    
Call RecosetMysql("select id, email from tblogin where email = '" & txtEmail.Text & "' and senha = '" & txtSenha.Text & "';")

If rs.EOF = False Then
    lblMsg.ForeColor = &HC000&
    lblMsg.Caption = "Acesso realizado com sucesso"
Else
    txtEmail.Text = ""
    txtSenha.Text = ""
    txtEmail.SetFocus
    lblMsg.ForeColor = &HFF&
    lblMsg.Caption = "Login não permitido"
End If

rs.Close
Set rs = Nothing

End Sub


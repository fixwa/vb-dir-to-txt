VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "pmn_DIR-TXT"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10935
   Icon            =   "pmn_DIR-TXT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3600
      TabIndex        =   5
      Top             =   300
      Width           =   2955
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Li&sto ahora CHAU me voy"
      Height          =   435
      Left            =   6960
      TabIndex        =   3
      Top             =   960
      Width           =   3795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&OK ahora abrí el pmnDIR.TXT"
      Height          =   435
      Left            =   180
      TabIndex        =   2
      Top             =   960
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Haceme el TXT ya!!!!!"
      Height          =   435
      Left            =   6960
      TabIndex        =   1
      Top             =   180
      Width           =   3735
   End
   Begin VB.Label Label2 
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
      Height          =   315
      Left            =   4080
      TabIndex        =   4
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecciona el disco que queres explorar:"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   300
      Width           =   3075
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

'If Drive1.Drive = "" Then GoTo a

If Text1.Text = "D:" Then GoTo Dp

If Text1.Text = "E:" Then GoTo Ep
If Text1.Text = "F:" Then GoTo Fp
If Text1.Text = "G:" Then GoTo Gp
If Text1.Text = "H:" Then GoTo Hp
If Text1.Text = "I:" Then GoTo Ip
If Text1.Text = "J:" Then GoTo Jp
If Text1.Text = "K:" Then GoTo Kp
If Text1.Text = "L:" Then GoTo Lp
If Text1.Text = "M:" Then GoTo Mp
If Text1.Text = "N:" Then GoTo Np
If Text1.Text = "O:" Then GoTo Op
If Text1.Text = "P:" Then GoTo Pp
If Text1.Text = "Q:" Then GoTo Qp
If Text1.Text = "R:" Then GoTo Rp
If Text1.Text = "S:" Then GoTo Sp
If Text1.Text = "T:" Then GoTo Tp
If Text1.Text = "U:" Then GoTo Up
If Text1.Text = "V:" Then GoTo Vp
If Text1.Text = "W:" Then GoTo Wp
If Text1.Text = "X:" Then GoTo Xp
If Text1.Text = "Y:" Then GoTo yp
If Text1.Text = "Z:" Then GoTo Zp
If Text1.Text = "C:" Then GoTo Cp

'ETIQUETAS
'Cp:
kk:
X = Shell("command.com /c dir C:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok

Dp:
X = Shell("command.com /c dir D:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Ep:
X = Shell("command.com /c dir E:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Fp:
X = Shell("command.com /c dir F:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Gp:
X = Shell("command.com /c dir G:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Hp:
X = Shell("command.com /c dir H:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Ip:
X = Shell("command.com /c dir I:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Jp:
X = Shell("command.com /c dir J:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Kp:
X = Shell("command.com /c dir K:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Lp:
X = Shell("command.com /c dir L:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Mp:
X = Shell("command.com /c dir M:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Np:
X = Shell("command.com /c dir N:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Op:
X = Shell("command.com /c dir O:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Pp:
X = Shell("command.com /c dir P:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Qp:
X = Shell("command.com /c dir Q:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Rp:
X = Shell("command.com /c dir R:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Sp:
X = Shell("command.com /c dir S:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Tp:
X = Shell("command.com /c dir T:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Up:
X = Shell("command.com /c dir U:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Vp:
X = Shell("command.com /c dir V:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Wp:
X = Shell("command.com /c dir W:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Xp:
X = Shell("command.com /c dir X:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
yp:
X = Shell("command.com /c dir Y:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok
Zp:
X = Shell("command.com /c dir Z:\*.* /s /oe >pmnDIR.TXT", vbHide)
GoTo allok

allok:
Label2.Caption = "LISTO !!!"

End Sub

Private Sub Command2_Click()
Dim y As String
y = Shell("notepad.exe pmnDIR.TXT", vbNormalFocus)

End Sub

Private Sub Command3_Click()
End
End Sub


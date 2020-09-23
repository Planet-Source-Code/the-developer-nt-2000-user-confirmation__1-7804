VERSION 5.00
Begin VB.Form frmpassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reviel "
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2595
   Icon            =   "Passwords.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDomain 
      Height          =   345
      Left            =   870
      TabIndex        =   3
      Top             =   735
      Width           =   1695
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   870
      TabIndex        =   1
      Top             =   15
      Width           =   1695
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   870
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Logon 
      Caption         =   "Log on"
      Default         =   -1  'True
      Height          =   345
      Left            =   840
      TabIndex        =   0
      Top             =   1110
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Domian"
      Height          =   195
      Left            =   15
      TabIndex        =   6
      Top             =   795
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   15
      TabIndex        =   5
      Top             =   405
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Name"
      Height          =   195
      Left            =   15
      TabIndex        =   4
      Top             =   45
      Width           =   795
   End
End
Attribute VB_Name = "frmpassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub logon_Click()
DoEvents
         
Logon.Enabled = False
txtDomain.Enabled = False
txtUserName.Enabled = False

Screen.MousePointer = vbHourglass

 DoEvents:        DoEvents
 
       If CheckPassword(txtDomain.Text, txtUserName.Text, txtPassword.Text, "") Then
             'Run your prgram now
                 MsgBox "Password confirmed."
                 Unload Me
                 End
                 'Call Main
                 
        Else
        
            MsgBox "Sorry incorrect password."
        End If


Screen.MousePointer = vbDefault
Logon.Enabled = True
txtDomain.Enabled = True
txtUserName.Enabled = True
Beep
End Sub




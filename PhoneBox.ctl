VERSION 5.00
Begin VB.UserControl PhoneBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      Begin VB.TextBox txtTEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   6
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox txtTEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   4
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtTEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   " -"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Parent2 
         BackColor       =   &H80000005&
         Caption         =   " ) -"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblParent1 
         BackColor       =   &H80000005&
         Caption         =   "("
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   135
      End
   End
End
Attribute VB_Name = "PhoneBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Sub txtTEL_KeyPress(Index As Integer, KeyAscii As Integer)
    'NUMBER ONLY'
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    'SWITCH FOCUS AFTER 3 CHAR OR SPACE'
    If Index < 2 Then
        If KeyAscii = 32 Then txtTEL(Index + 1).SetFocus
        
        If Len(txtTEL(Index)) > 1 Then
            txtTEL(Index + 1).SetFocus
        End If
    Else
        txtTEL(Index).MaxLength = 4
    End If
End Sub

'RETURN VAL'
Public Property Get Text() As String
    Text = "(" & txtTEL(0).Text & ")" & "-" & txtTEL(1).Text & "-" & txtTEL(2).Text
End Property

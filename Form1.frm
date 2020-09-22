VERSION 5.00
Object = "{4225190B-AB4B-40F0-A4B5-BFE3377A69B8}#2.0#0"; "d1rtyBOX.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin d1rtyBOX.IPbox IPbox1 
      Height          =   315
      Left            =   1020
      TabIndex        =   2
      Top             =   420
      Width           =   3495
      _extentx        =   6165
      _extenty        =   556
      couleurfond     =   -2147483633
      backcolor       =   -2147483643
      byte1           =   "151"
      byte1           =   "151"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1380
      TabIndex        =   1
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Afficher le IP"
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    MsgBox IPbox1.AdresseIP
    MsgBox IPbox1.Byte0
    'MsgBox PhoneBox1.Text
End Sub

Private Sub Command2_Click()
    MsgBox IPbox1.Byte0
    IPbox1.Byte0_FillColor = vbRed
End Sub

VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   2400
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   14
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Caption         =   "General"
      Height          =   975
      Left            =   0
      TabIndex        =   11
      Top             =   1320
      Width           =   5175
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2820
         TabIndex        =   16
         Text            =   "80"
         Top             =   240
         Width           =   555
      End
      Begin VB.CheckBox chkHost 
         Caption         =   "Resolve Host"
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   300
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.CheckBox chkAlive 
         Caption         =   "Reply Only"
         Height          =   195
         Left            =   300
         TabIndex        =   12
         Top             =   600
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Filter Ports "
         Height          =   195
         Left            =   1920
         TabIndex        =   15
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Port Scan"
      Height          =   1275
      Left            =   2640
      TabIndex        =   7
      Top             =   0
      Width           =   2535
      Begin VB.CheckBox Check1 
         Caption         =   "Scan Alive Only"
         Enabled         =   0   'False
         Height          =   195
         Left            =   480
         TabIndex        =   17
         Top             =   840
         Value           =   1  'Checked
         Width           =   1755
      End
      Begin VB.TextBox txtPrtTimeOut 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         TabIndex        =   8
         Text            =   "300"
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Timeout"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "ms"
         Height          =   195
         Left            =   1860
         TabIndex        =   9
         Top             =   420
         Width           =   255
      End
   End
   Begin VB.Frame frmPing 
      Caption         =   "Ping"
      Height          =   1275
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2595
      Begin VB.TextBox txtpngTimeOut 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         TabIndex        =   2
         Text            =   "300"
         Top             =   360
         Width           =   795
      End
      Begin VB.TextBox txtpngTTL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         TabIndex        =   1
         Text            =   "150"
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Timeout"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "ms"
         Height          =   195
         Left            =   1860
         TabIndex        =   5
         Top             =   420
         Width           =   255
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "TTL"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label8 
         Caption         =   "hops"
         Height          =   195
         Left            =   1860
         TabIndex        =   3
         Top             =   840
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Dim tmpHost As Boolean, tmpReply As Boolean
    
    If chkHost.Value = 1 Then
        tmpHost = True
    Else
        tmpHost = False
    End If
    
    If chkAlive.Value = 0 Then
        tmpReply = True
    Else
        tmpReply = False
    End If
    
    SaveOptions CStr(tmpHost), CStr(tmpReply), CStr(txtpngTTL.Text), CStr(txtpngTimeOut.Text), CStr(3), CStr(txtPrtTimeOut.Text)
    Unload Me
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    GetOptions
    
    
    txtpngTimeOut.Text = pngTimeOut
    txtpngTTL.Text = pngTTL
    txtPrtTimeOut.Text = prtTimeOut
    
    If genHost = True Then
        chkHost.Value = 1
    Else
        chkHost.Value = 0
    End If
    
    If genReply = True Then
        chkAlive.Value = 0
    Else
        chkAlive.Value = 1
    End If

End Sub

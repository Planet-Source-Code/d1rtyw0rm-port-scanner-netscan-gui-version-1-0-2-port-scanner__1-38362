VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4225190B-AB4B-40F0-A4B5-BFE3377A69B8}#2.0#0"; "d1rtyBOX.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "P.W.A - Network Scanner"
   ClientHeight    =   7995
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   10095
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   533
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   673
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDial 
      Left            =   9480
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   885
      Left            =   6660
      Top             =   1320
   End
   Begin MSComctlLib.ImageList imgAnim 
      Left            =   6720
      Top             =   1740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D96
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTools 
      Left            =   9480
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26EA
            Key             =   "Launch"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8EDC
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EC3E
            Key             =   "SaveAs"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14860
            Key             =   "Clear"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17012
            Key             =   "Option"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":197C4
            Key             =   "Help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   1535
      ButtonWidth     =   1773
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "imgTools"
      HotImageList    =   "imgTools"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Launch Scan"
            Key             =   "tlsLaunch"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Abort Scan"
            Key             =   "tlsAbort"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear Result"
            Key             =   "tlsClear"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save Result"
            Key             =   "tlsSave"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            Key             =   "tlsOptions"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "tlsHelp"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrConnected 
      Left            =   9480
      Top             =   4320
   End
   Begin MSWinsockLib.Winsock sckScan 
      Left            =   9480
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgTree 
      Left            =   9480
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F3EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20D80
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22B8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22FDC
            Key             =   "ftp"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2338E
            Key             =   "http"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":236E0
            Key             =   "smtp"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23FBA
            Key             =   "mssql"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2430C
            Key             =   "https"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24660
            Key             =   "else"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Scan Type"
      Height          =   1815
      Left            =   0
      TabIndex        =   13
      Top             =   900
      Width           =   4635
      Begin d1rtyBOX.IPbox IPboxFROM 
         Height          =   315
         Left            =   420
         TabIndex        =   1
         Top             =   840
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         CouleurFond     =   -2147483633
         BackColor       =   -2147483643
      End
      Begin VB.OptionButton optSingle 
         Caption         =   "Single IP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2340
         TabIndex        =   17
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton optRange 
         Caption         =   "Range Scan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
      Begin d1rtyBOX.IPbox IPboxTO 
         Height          =   255
         Left            =   420
         TabIndex        =   2
         Top             =   1440
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   450
         CouleurFond     =   -2147483633
         BackColor       =   -2147483643
      End
      Begin d1rtyBOX.IPbox IPboxSingle 
         Height          =   315
         Left            =   2700
         TabIndex        =   4
         Top             =   840
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         CouleurFond     =   -2147483633
         Enabled         =   0   'False
         BackColor       =   -2147483643
      End
      Begin VB.Label Label2 
         Caption         =   "IP Address"
         Height          =   195
         Left            =   2700
         TabIndex        =   18
         Top             =   660
         Width           =   795
      End
      Begin VB.Label Label6 
         Caption         =   "To"
         Height          =   195
         Left            =   420
         TabIndex        =   16
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "From"
         Height          =   195
         Left            =   420
         TabIndex        =   15
         Top             =   660
         Width           =   1515
      End
   End
   Begin VB.Frame frmScanOption 
      Caption         =   "PortScan Configuration"
      Height          =   1830
      Left            =   4680
      TabIndex        =   12
      Top             =   900
      Width           =   3255
      Begin VB.TextBox txtManual 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   420
         TabIndex        =   10
         Text            =   "7:21:23:80:137:138:139"
         Top             =   1500
         Width           =   2715
      End
      Begin VB.OptionButton optManual 
         Caption         =   "Define Manualy"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   1260
         Width           =   1515
      End
      Begin VB.TextBox txtPrtTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1020
         TabIndex        =   8
         Text            =   "65535"
         Top             =   900
         Width           =   555
      End
      Begin VB.TextBox txtPrtFrom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   420
         TabIndex        =   7
         Text            =   "1"
         Top             =   900
         Width           =   555
      End
      Begin VB.OptionButton OptFromTo 
         Caption         =   "Define From - To"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   660
         Width           =   1695
      End
      Begin VB.OptionButton optPingOnly 
         Caption         =   "Ping Only"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1875
      End
      Begin VB.Image imgPort 
         Height          =   720
         Left            =   2460
         Picture         =   "frmMain.frx":24948
         Stretch         =   -1  'True
         Top             =   180
         Width           =   780
      End
   End
   Begin VB.Timer tmrTimeOut 
      Left            =   9480
      Top             =   5400
   End
   Begin MSComctlLib.ProgressBar pbScan 
      Height          =   255
      Left            =   90
      TabIndex        =   11
      Top             =   7425
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar sts 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7740
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   9604
            Text            =   "Waiting ..."
            TextSave        =   "Waiting ..."
            Object.ToolTipText     =   "Current Action"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Text            =   "000.000.000.000"
            TextSave        =   "000.000.000.000"
            Object.ToolTipText     =   "IP Adress"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Text            =   "N/A"
            TextSave        =   "N/A"
            Object.ToolTipText     =   "Scanning Port"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Text            =   "PC Founded : N/A"
            TextSave        =   "PC Founded : N/A"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvResult 
      Height          =   4605
      Left            =   0
      TabIndex        =   19
      Top             =   2760
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8123
      _Version        =   393217
      Indentation     =   661
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "imgTree"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   7980
      MousePointer    =   2  'Cross
      Picture         =   "frmMain.frx":24C54
      Stretch         =   -1  'True
      Top             =   900
      Width           =   2115
   End
   Begin VB.Menu mnuResult 
      Caption         =   "Menu TreeView Result Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuIPtoCLIP 
         Caption         =   "Copy IP to clipboard"
      End
      Begin VB.Menu mnusep903 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearAll 
         Caption         =   "Clear All"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Selected"
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Filter"
      End
      Begin VB.Menu Sep837 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColapse 
         Caption         =   "Collapse All"
      End
      Begin VB.Menu mnuExpand 
         Caption         =   "Expand All"
      End
      Begin VB.Menu sep324 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConnTo 
         Caption         =   "Connect To"
         Begin VB.Menu mnuFTP 
            Caption         =   "FTP"
         End
         Begin VB.Menu mnuTelnet 
            Caption         =   "TELNET"
         End
         Begin VB.Menu mnuHTTP 
            Caption         =   "HTTP"
         End
         Begin VB.Menu mnuThisPort 
            Caption         =   "This Port"
            Begin VB.Menu mnuThisFTP 
               Caption         =   "Via FTP"
            End
            Begin VB.Menu mnuThisTelnet 
               Caption         =   "Via TELNET"
            End
            Begin VB.Menu mnuThisHttp 
               Caption         =   "Via HTTP"
            End
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objPing As InternetTools.clsPing 'OBJECT CONTAIN PING & RESOLVE HOST FUNCTION'
Dim Go As Boolean, Data As String

'*************************************************************************************'
'                               MENU FILE CODING SECTION                              '
'*************************************************************************************'

Private Sub mnuClearAll_Click()
    tvResult.Nodes.Clear
End Sub

Private Sub mnuColapse_Click()
Dim x As Long
    For x = 1 To tvResult.Nodes.Count
        tvResult.Nodes(x).Expanded = False
    Next x
End Sub

Private Sub mnuDelete_Click()
    tvResult.Nodes.Remove (tvResult.SelectedItem.Index)
End Sub

Private Sub mnuExpand_Click()
    For x = 1 To tvResult.Nodes.Count
        tvResult.Nodes(x).Expanded = True
    Next x
End Sub

Private Sub mnuFilter_Click()
Dim Marked() As Integer
'SO FUCKED UP ?!?!?!'
    ReDim Marked(1)
    
    For x = 1 To tvResult.Nodes.Count Step 1
        If Left(tvResult.Nodes(x).Key, 4) = "Port" Then
            If Mid(tvResult.Nodes(x).Key, InStr(tvResult.Nodes(x).Key, "-") + 1, Len(tvResult.Nodes(x).Key)) = 80 Then
                Marked(UBound(Marked)) = tvResult.Nodes(x).Parent.Index
                ReDim Preserve Marked(UBound(Marked) + 1)
            End If
        End If
    Next x
    
Reset:
    For x = 1 To tvResult.Nodes.Count Step 1
        If Left(tvResult.Nodes(x).Key, 3) = "Box" Then
            For y = 1 To UBound(Marked) Step 1
                If tvResult.Nodes(x).Index = Marked(y) Then
                    GoTo NextOne
                End If
                
                tvResult.Nodes.Remove (x)
                tvResult.Refresh
            Next y
            GoTo Reset
        End If
NextOne:
    Next x
End Sub

Private Sub mnuFTP_Click()
Dim strIP As String
On Error GoTo findIT
    strIP = Mid(tvResult.SelectedItem.Key, 4, Len(tvResult.SelectedItem.Key))
    If tvResult.SelectedItem.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Key))
    End If
    
    If tvResult.SelectedItem.Parent.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Parent.Key))
    End If
    
    If tvResult.SelectedItem.Parent.Parent.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Parent.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Parent.Parent.Key))
    End If

    Exit Sub
findIT:
    ShellExecute Me.hwnd, vbNullString, "ftp://" & strIP & ":21", vbNullString, "C:\", 1
    Exit Sub
End Sub

Private Sub mnuHTTP_Click()
Dim strIP As String
On Error GoTo findIT
    strIP = Mid(tvResult.SelectedItem.Key, 4, Len(tvResult.SelectedItem.Key))
    If tvResult.SelectedItem.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Key))
    End If
    
    If tvResult.SelectedItem.Parent.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Parent.Key))
    End If
    
    If tvResult.SelectedItem.Parent.Parent.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Parent.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Parent.Parent.Key))
    End If

    Exit Sub
findIT:
    ShellExecute Me.hwnd, vbNullString, "http://" & strIP, vbNullString, "C:\", 1
    Exit Sub
End Sub

Private Sub mnuIPtoCLIP_Click()
Dim strIP As String
On Error GoTo findIT
    strIP = Mid(tvResult.SelectedItem.Key, 4, Len(tvResult.SelectedItem.Key))
    If tvResult.SelectedItem.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Key))
    End If
    
    If tvResult.SelectedItem.Parent.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Parent.Key))
    End If
    
    If tvResult.SelectedItem.Parent.Parent.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Parent.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Parent.Parent.Key))
    End If
    
    Clipboard.SetText strIP
    Exit Sub
findIT:
    Clipboard.SetText strIP
    Exit Sub
End Sub

Private Sub mnuQuit_Click()
    End
End Sub

Private Sub mnuTelnet_Click()
Dim strIP As String
On Error GoTo findIT
    strIP = Mid(tvResult.SelectedItem.Key, 4, Len(tvResult.SelectedItem.Key))
    If tvResult.SelectedItem.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Key))
    End If
    
    If tvResult.SelectedItem.Parent.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Parent.Key))
    End If
    
    If tvResult.SelectedItem.Parent.Parent.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Parent.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Parent.Parent.Key))
    End If

    Exit Sub
findIT:
    Shell "telnet " & strIP & " " & "23", vbNormalFocus
    Exit Sub
End Sub

Private Sub mnuThisFTP_Click()
Dim strPort As String
On Error GoTo findIT

    strIP = Mid(tvResult.SelectedItem.Key, 4, Len(tvResult.SelectedItem.Key))
    strPort = Left(tvResult.SelectedItem.Text, InStr(tvResult.SelectedItem.Text, "/") - 1)
    
    If tvResult.SelectedItem.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Key))
    End If
    
    If tvResult.SelectedItem.Parent.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Parent.Key))
    End If
    
    If tvResult.SelectedItem.Parent.Parent.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Parent.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Parent.Parent.Key))
    End If

    Exit Sub
findIT:
    ShellExecute Me.hwnd, vbNullString, "ftp://" & strIP & ":" & strPort, vbNullString, "C:\", 1
    Exit Sub
End Sub

Private Sub mnuThisHttp_Click()
Dim strPort As String
On Error GoTo findIT

    strIP = Mid(tvResult.SelectedItem.Key, 4, Len(tvResult.SelectedItem.Key))
    strPort = Left(tvResult.SelectedItem.Text, InStr(tvResult.SelectedItem.Text, "/") - 1)
    
    If tvResult.SelectedItem.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Key))
    End If
    
    If tvResult.SelectedItem.Parent.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Parent.Key))
    End If
    
    If tvResult.SelectedItem.Parent.Parent.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Parent.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Parent.Parent.Key))
    End If

    Exit Sub
findIT:
    ShellExecute Me.hwnd, vbNullString, "Http://" & strIP & ":" & strPort, vbNullString, "C:\", 1
    Exit Sub
End Sub

Private Sub mnuThisTelnet_Click()
Dim strPort As String
On Error GoTo findIT

    strIP = Mid(tvResult.SelectedItem.Key, 4, Len(tvResult.SelectedItem.Key))
    strPort = Left(tvResult.SelectedItem.Text, InStr(tvResult.SelectedItem.Text, "/") - 1)
    
    If tvResult.SelectedItem.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Key))
    End If
    
    If tvResult.SelectedItem.Parent.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Parent.Key))
    End If
    
    If tvResult.SelectedItem.Parent.Parent.Parent.Key <> "" Then
        strIP = Mid(tvResult.SelectedItem.Parent.Parent.Parent.Key, 4, Len(tvResult.SelectedItem.Parent.Parent.Parent.Key))
    End If

    Exit Sub
findIT:
    Shell "telnet " & strIP & " " & strPort, vbNormalFocus
    Exit Sub

End Sub

'*************************************************************************************'
'                               END OF MENU CODING SECTION                            '
'*************************************************************************************'

Private Sub Form_Load()
Dim tmpStr As String * 300, tmpLen As String

    Go = False
    tmpLen = GetPrivateProfileString("PORT", "ManualString", "7:21:23:25:79:80:135:137:139:443:513:1080:1433:1525:1527:1745:3306:3351:5631:6000:12345:31337:54320:54321", tmpStr, Len(tmpStr), lpFileName)
        txtManual.Text = Trim(tmpStr)
End Sub

Private Sub Form_Resize()
    ResizeForm Me 'RESIZE ALL CONTROL TO FIT WITH THE SCREEN'
End Sub

Private Sub Image1_Click()
    'GOTO PWA WEB PAGE'
    ShellExecute Me.hwnd, vbNullString, "http://www.pwa.ca.tc", vbNullString, "C:\", 1
End Sub


Private Sub OptFromTo_Click()
    txtPrtFrom.Enabled = True
    txtPrtTo.Enabled = True
    txtManual.Enabled = False
End Sub

Private Sub optManual_Click()
    txtPrtFrom.Enabled = False
    txtPrtTo.Enabled = False
    txtManual.Enabled = True
End Sub

Private Sub optPingOnly_Click(Index As Integer)
    txtPrtFrom.Enabled = False
    txtPrtTo.Enabled = False
    txtManual.Enabled = False
End Sub

Private Sub optRange_Click()
    IPboxSingle.Enabled = False
    IPboxFROM.Enabled = True
    IPboxTO.Enabled = True
End Sub

Private Sub optSingle_Click()
    IPboxSingle.Enabled = True
    IPboxFROM.Enabled = False
    IPboxTO.Enabled = False
End Sub


'*************************************************************************************'
'   Function    :   Scan()                                                            '
'   Description :   Try to connect on a remote box port using winsock and retreive    '
'                   what services is running on.                                      '
'   Paramater   :                                                                     '
'                   HostName    :   Target IP                                         '
'                   Port        :   Port to Scan                                      '
'                   TimeOut     :   Time before stop of port reply                    '
'*************************************************************************************'
Public Function Scan(ByVal HostName As String, ByVal Port As Long, _
    ByVal TIMEOUTms As Long)
    
    'SET WINSOCK INFORMATION'
    sckScan.Close
    sckScan.RemoteHost = HostName
    sckScan.RemotePort = Port
    
    'PRINT STATUS TO USER'
    sts.Panels(1).Text = "Port Scanning"
    sts.Panels(2).Text = HostName
    sts.Panels(3).Text = Port
    
    'CONNECT TO SOCKET'
    sckScan.Connect
    
    'IF CONNECTIONTIME > TIMEOUTms, inactive port'
    tmrTimeOut.Interval = TIMEOUTms
    tmrTimeOut.Enabled = True
    
    
    
End Function

'*************************************************************************************'
' Make the animation when scanning is active                                          '
'*************************************************************************************'
Private Sub Timer1_Timer()
Static x As Integer
    If x >= 4 Then
        x = 1
    Else
        x = x + 1
    End If
    
    imgPort.Picture = imgAnim.ListImages(x).Picture
    DoEvents
End Sub

'*************************************************************************************'
'STOP TRYING CONNECT, PORT IS INNACTIVE                                               '
'*************************************************************************************'
Private Sub tmrTimeOut_Timer()
    sckScan.Close
    Data = ""
    tmrTimeOut.Enabled = False
    Go = True
End Sub

'*************************************************************************************'
' GET DATA FROM THE PORT TO DEFINE WHAT SERVICE IS RUNNING ON                         '
'*************************************************************************************'
Private Sub sckScan_DataArrival(ByVal bytesTotal As Long)
    sckScan.GetData Data, vbString
      
    If InStr(1, Data, vbLf) <> 0 Then
        'ResolvePort Data
        sckScan.Close
    End If
End Sub

'*************************************************************************************'
' ON PORT CONNECT, SEND TO PORT A STUPID STING TO SEE WHAT WILL BE THE REPLY          '
'*************************************************************************************'
Private Sub sckScan_Connect()
    tmrTimeOut.Enabled = False

    sckScan.SendData "GET index.html" & vbCrLf & "USER fake" & vbCrLf & "FINGER fake" & vbCrLf
    tmrConnected.Interval = 2000
    tmrConnected.Enabled = True
End Sub

'*************************************************************************************'
' PRINT RESULT ON SCANNED PC PORT                                                     '
'*************************************************************************************'
Private Sub tmrConnected_Timer()
Dim tmpStr As String
Static PortKey As Long
    
    Select Case sckScan.RemotePort
        Case 21:
                tvResult.Nodes.Add "Box" & sckScan.RemoteHost, tvwChild, "Port" & sckScan.RemoteHost & "-" & sckScan.RemotePort, ResolvePort(sckScan.RemotePort), "ftp"
        Case 25:
                tvResult.Nodes.Add "Box" & sckScan.RemoteHost, tvwChild, "Port" & sckScan.RemoteHost & "-" & sckScan.RemotePort, ResolvePort(sckScan.RemotePort), "smtp"
        Case 80:
                tvResult.Nodes.Add "Box" & sckScan.RemoteHost, tvwChild, "Port" & sckScan.RemoteHost & "-" & sckScan.RemotePort, ResolvePort(sckScan.RemotePort), "http"
        Case 443:
                tvResult.Nodes.Add "Box" & sckScan.RemoteHost, tvwChild, "Port" & sckScan.RemoteHost & "-" & sckScan.RemotePort, ResolvePort(sckScan.RemotePort), "https"
        Case 1433:
                tvResult.Nodes.Add "Box" & sckScan.RemoteHost, tvwChild, "Port" & sckScan.RemoteHost & "-" & sckScan.RemotePort, ResolvePort(sckScan.RemotePort), "mssql"
        Case Else:
                tvResult.Nodes.Add "Box" & sckScan.RemoteHost, tvwChild, "Port" & sckScan.RemoteHost & "-" & sckScan.RemotePort, ResolvePort(sckScan.RemotePort), "else"
    End Select
    
    If Data <> "" Then
        If InStr(Data, vbCrLf) <> 0 Then
            Do Until InStr(Data, vbLf) = 0
                tmpStr = Left(Data, InStr(Data, vbCrLf) - 1)
                If tmpStr <> "" Then
                    PortKey = PortKey + 1
                    tvResult.Nodes.Add "Port" & sckScan.RemoteHost & "-" & sckScan.RemotePort, tvwChild, "Info" & "+" & PortKey & "-" & sckScan.RemoteHost & ":" & sckScan.RemotePort, CStr(tmpStr), 3
                End If
                
                Data = Right(Data, (Len(Data) - Len(tmpStr)) - 2)
            Loop
        End If
    End If
    tmrConnected.Enabled = False
    Go = True
End Sub

'*************************************************************************************'
' SELECT AN ACTION TO DO IN THE TOOLBAR MENU                                          '
'                                                                                     '
' LAUNCH - ABORT - CLEAR LIST - SAVE - OPTIONS - HELP                                 '
'*************************************************************************************'
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Cmptr As Integer, Cmptr1 As Integer
Dim Replyms As Long, boxNum As Long, tmpPB As Long
Dim tmpReply() As String, strPort As String, tmpIP As String
Dim CurrentBox As String, CurrentPort As Long
Dim x As Integer, y As Integer, z As Integer

Const OK = 0, RangeToLong = 1, BadRange = 2

Set objPing = New InternetTools.clsPing
GetOptions

Select Case Button.Key
    'LAUNCH SCAN'
    Case "tlsLaunch":
        'SINGLE SCAN'
        If optSingle.Value = True Then
            tmpIP = CStr(IPboxSingle.Byte0) & "." & CStr(IPboxSingle.Byte1) & "." & CStr(IPboxSingle.byte2) & "." & CStr(IPboxSingle.byte3)
            Replyms = objPing.Ping(tmpIP, pngTTL, pngTimeOut)
                If Replyms <> -666 Then
                    boxNum = boxNum + 1
                    sts.Panels(4).Text = "PC Founded : " & boxNum
                    If genHost = True Then
                        tvResult.Nodes.Add , , "Box" & tmpIP, CStr(tmpIP) & " (" & objPing.ResolveHostname(tmpIP) & ") " & Replyms & "ms", 1
                        tvResult.Refresh
                        DoEvents
                    Else
                        tvResult.Nodes.Add , , "Box" & tmpIP, "Reply from " & tmpIP & " recieved after " & Replyms & "ms", 1
                        tvResult.Refresh
                        DoEvents
                    End If
                Else
                    If genReply = True Then
                        tvResult.Nodes.Add , , "Fail" & tmpIP, CStr(tmpIP) & " FAILLED !", 2
                        tvResult.Refresh
                        DoEvents
                    End If
                End If
                
            'PORTSCAN
            If OptFromTo.Value = True Then
                pbScan.Min = CLng(txtPrtFrom.Text)
                pbScan.Max = CLng(txtPrtTo.Text)
                For y = CLng(txtPrtFrom.Text) To CLng(txtPrtTo.Text) Step 1
                    pbScan.Value = y
                    Scan tmpIP, y, prtTimeOut
                    Do Until Go = True
                        DoEvents
                    Loop
                    Go = False
                Next y
        End If
        
        If optManual.Value = True Then
            Cmptr = 0
            strPort = txtManual.Text
            
            Do While InStr(strPort, ":")
                Cmptr = Cmptr + 1
                strtmp = Left(strPort, InStr(strPort, ":") - 1)
                strPort = Right(strPort, (Len(strPort) - Len(strtmp) - 1))
                    
                Scan tmpIP, CLng(strtmp), prtTimeOut
                Do Until Go = True
                    DoEvents
                Loop
                Go = False
            Loop
        End If
        Timer1.Enabled = False
        imgPort.Picture = imgAnim.ListImages(1).Picture
        sts.Panels(1).Text = "Scan Complete !"
        sts.Panels(2).Text = "000.000.000.000"
        sts.Panels(3).Text = "N/A"
        sts.Panels(4).Text = "PC Founded : " & boxNum
        boxNum = 0
        Set objPing = Nothing
            Exit Sub
        End If
        
        'RANGE SCAN'
        Select Case ValidateScan()
            Case OK:
                ReDim tmpReply(0)
                sts.Panels(1).Text = "Scan in progress ..."
                Timer1.Enabled = True
                
            Case RangeToLong:
                MsgBox "Range to scan is too long, restrict your scan to B range !", vbInformation, "NetScan"
                Exit Sub
                
            Case BadRange:
                MsgBox "Invalid scanning range !", vbCritical, "Error"
                Exit Sub
        End Select
        
        'AJUST PROGRESS BAR'
        If IPboxFROM.byte2 = IPboxTO.byte2 Then
            boxNum = (IPboxTO.byte3 - IPboxFROM.byte3) + 1
            pbScan.Max = boxNum
            pbScan.Min = 0
        Else
            Cmptr = IPboxTO.byte2 - IPboxFROM.byte2
            
            If Cmptr < 3 Then
                Cmptr1 = 255 - IPboxFROM.byte3
                boxNum = Cmptr * Cmptr1 + IPboxTO.byte3
            Else
                Cmptr1 = 255 - IPboxFROM.byte3
                boxNum = ((((IPboxTO.byte2 - IPboxFROM.byte2) - 1) * 255) + IPboxTO.byte3 + Cmptr1)
            End If
            
            pbScan.Max = boxNum
            pbScan.Min = 0
        End If
        
        tmpPB = 0
        boxNum = 0
        
        Dim byte2 As Integer, byte3 As Integer
        byte2 = IPboxFROM.byte2
        byte3 = IPboxFROM.byte3
        
        Do While (byte2 <> CInt(IPboxTO.byte2)) Or (byte3 <> CInt(IPboxTO.byte3) + 1)
            
            tmpIP = CStr(IPboxFROM.Byte0) & "." & CStr(IPboxFROM.Byte1) & "." & byte2 & "." & byte3
                
                'START THAT SHIT'
                DoEvents
                tmpPB = tmpPB + 1
                On Error Resume Next
                pbScan.Value = tmpPB
                sts.Panels(1).Text = "Pinging"
                sts.Panels(2).Text = tmpIP
                sts.Panels(3).Text = "N/A"
                Replyms = objPing.Ping(tmpIP, pngTTL, pngTimeOut)
                If Replyms <> -666 Then
                    boxNum = boxNum + 1
                    sts.Panels(4).Text = "PC Founded : " & boxNum
                    If genHost = True Then
                        tvResult.Nodes.Add , , "Box" & tmpIP, CStr(tmpIP) & " (" & objPing.ResolveHostname(tmpIP) & ") " & Replyms & "ms", 1
                        ReDim Preserve tmpReply(UBound(tmpReply) + 1)
                        tmpReply(UBound(tmpReply)) = tmpIP
                        tvResult.Refresh
                        DoEvents
                    Else
                        tvResult.Nodes.Add , , "Box" & tmpIP, "Reply from " & tmpIP & " recieved after " & Replyms & "ms", 1
                        ReDim Preserve tmpReply(UBound(tmpReply) + 1)
                        tmpReply(UBound(tmpReply)) = tmpIP
                        tvResult.Refresh
                        DoEvents
                    End If
                Else
                    If genReply = True Then
                        tvResult.Nodes.Add , , "Fail" & tmpIP, CStr(tmpIP) & " FAILLED !", 2
                        tvResult.Refresh
                        DoEvents
                    End If
                End If
                
            'ONE MORE TIME'
                If byte2 >= IPboxTO.byte2 Then
                    If byte3 = IPboxTO.byte3 Then
                        Exit Do
                    End If
                End If
                
                If byte3 = 255 Then
                    byte3 = 1
                    byte2 = byte2 + 1
                Else
                    byte3 = byte3 + 1
                End If
            Loop
        
        
        'SCAN PORT'
        If OptFromTo.Value = True Then
            Cmptr = CLng(txtPrtTo.Text) - CLng(txtPrtFrom)

            pbScan.Min = 0
            pbScan.Max = (Cmptr * UBound(tmpReply)) + 4
            Cmptr = 0
            For x = 1 To UBound(tmpReply)
                For y = CLng(txtPrtFrom.Text) To CLng(txtPrtTo.Text) Step 1
                    Scan tmpReply(x), y, prtTimeOut
                    Cmptr = Cmptr + 1
                    pbScan.Value = Cmptr
                    Do Until Go = True
                        DoEvents
                    Loop
                    Go = False
                Next y
            Next x
        End If
        
        If optManual.Value = True Then
            Cmptr = 0
            strPort = txtManual.Text
            
            Do While InStr(strPort, ":")
                Cmptr = Cmptr + 1
                strtmp = Left(strPort, InStr(strPort, ":") - 1)
                strPort = Right(strPort, (Len(strPort) - Len(strtmp) - 1))
            Loop
            
            pbScan.Min = 0
            pbScan.Max = UBound(tmpReply) * Cmptr
            
            Cmptr = 0
            
            For x = 1 To UBound(tmpReply)
                strPort = txtManual.Text
                    Do While InStr(strPort, ":")
                        Cmptr = Cmptr + 1
                        pbScan.Value = Cmptr
                        strtmp = Left(strPort, InStr(strPort, ":") - 1)
                        strPort = Right(strPort, (Len(strPort) - Len(strtmp) - 1))
                        
                        Scan tmpReply(x), CLng(strtmp), prtTimeOut
                        Do Until Go = True
                            DoEvents
                        Loop
                        Go = False
                    Loop
            Next x
        End If
        Timer1.Enabled = False
        imgPort.Picture = imgAnim.ListImages(1).Picture
        sts.Panels(1).Text = "Scan Complete !"
        sts.Panels(2).Text = "000.000.000.000"
        sts.Panels(3).Text = "N/A"
        sts.Panels(4).Text = "PC Founded : " & boxNum
        boxNum = 0
        Set objPing = Nothing
    
    Case "tlsAbort":
        sckScan.Close
        Set objPing = Nothing
        
    Case "tlsClear":
        tvResult.Nodes.Clear
    
    Case "tlsSave":
    CDial.Filter = "HTM Files (*.htm)|*.htm|" & " HTML Files (*.html)|*.html|"
    CDial.ShowSave
    If CDial.FileName <> "" Then
        'GENERATE HTML REPORT'
        Open CDial.FileName For Output As #1
            Call CopyIco(CDial.FileName)
            
            Print #1, "<HTML><HEAD><TITLE>NetScan Report For " & IPboxFROM.AdresseIP & " - " & IPboxTO.AdresseIP & "</TITLE></HEAD><BODY>"
            Print #1, "<IMG SRC=Title1.bmp>"
            For x = 1 To tvResult.Nodes.Count Step 1
                If Left(tvResult.Nodes(x).Key, 3) = "Box" Then
                    Print #1, "<BR><HR><BR><B><IMG SRC=Comp.bmp>" & tvResult.Nodes(x).Text & "</B><BR>"
                    CurrentBox = Mid(tvResult.Nodes(x).Key, 4, Len(tvResult.Nodes(x)))
                    
                    For y = 1 To tvResult.Nodes.Count Step 1
                        If Left(tvResult.Nodes(y).Key, 4) = "Port" Then
                            If Mid(tvResult.Nodes(y).Key, 5, InStr(tvResult.Nodes(y).Key, "-") - 5) = CurrentBox Then
                                CurrentPort = Mid(tvResult.Nodes(y).Key, InStr(tvResult.Nodes(y).Key, "-") + 1, Len(tvResult.Nodes(y).Key))
                                Select Case CurrentPort
                                        Case 21:
                                            Print #1, "<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<U><IMG SRC=FTP.bmp>" & tvResult.Nodes(y).Text & "</U><BR>"
                                        Case 25:
                                            Print #1, "<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<U><IMG SRC=SMTP.bmp>" & tvResult.Nodes(y).Text & "</U><BR>"
                                        Case 80:
                                            Print #1, "<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<U><IMG SRC=HTTP.bmp>" & tvResult.Nodes(y).Text & "</U><BR>"
                                        Case 443:
                                            Print #1, "<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<U><IMG SRC=HTTPS.bmp>" & tvResult.Nodes(y).Text & "</U><BR>"
                                        Case 1433:
                                            Print #1, "<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<U><IMG SRC=SQL.bmp>" & tvResult.Nodes(y).Text & "</U><BR>"
                                        Case Else:
                                            Print #1, "<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<U><IMG SRC=Ports.bmp>" & tvResult.Nodes(y).Text & "</U><BR>"
                                End Select
                                
                                For z = 1 To tvResult.Nodes.Count Step 1
                                    If Left(tvResult.Nodes(z).Key, 4) = "Info" Then
                                        If Mid(tvResult.Nodes(z).Key, InStr(tvResult.Nodes(z).Key, "-") + 1, InStr(tvResult.Nodes(z).Key, ":") - 8) = CurrentBox Then
                                            If Mid(tvResult.Nodes(z).Key, InStr(tvResult.Nodes(z).Key, ":") + 1, Len(tvResult.Nodes(z).Key)) = CurrentPort Then
                                                Print #1, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I><IMG SRC=Orb.bmp>" & tvResult.Nodes(z).Text & "</I><BR>"
                                            End If
                                        End If
                                    End If
                                Next z
                            End If
                        End If
                    Next y
                End If
            Next x
            
            Print #1, "</BODY></HTML>"
        Close #1
        
        MsgBox "HTML Report has been successfuly created !", vbInformation, "NetScan Report"
    End If

    Case "tlsOptions":
        frmOptions.Show vbModal, frmMain
        
    Case "tlsHelp":
        MsgBox "Call now to our free help line" & vbCrLf & "-=         1-976-6725          =-", vbInformation, "FREE HELP LINE"
        
End Select

End Sub

'NUM ONLY'
Private Sub txtTimeOut_KeyPress(KeyAscii As Integer)
    If KeyAscii > 57 Or KeyAscii < 48 Then KeyAscii = 0
End Sub

Private Function ValidateScan() As Integer
Const OK = 0, RangeToLong = 1, BadRange = 2
    With frmMain
        If .IPboxFROM.Byte0 <> .IPboxTO.Byte0 Then
            ValidateScan = RangeToLong
            Exit Function
            
        ElseIf .IPboxFROM.Byte1 <> .IPboxTO.Byte1 Then
            ValidateScan = RangeToLong
            Exit Function
        
        ElseIf .IPboxFROM.byte2 > .IPboxTO.byte2 Then
            ValidateScan = BadRange
            Exit Function
        
        ElseIf CLng(.IPboxFROM.byte3) > CLng(.IPboxTO.byte3) Then
            If CLng(.IPboxFROM.byte2) = CLng(.IPboxTO.byte2) Then
                ValidateScan = BadRange
            End If
            Exit Function
        End If
    End With
    
    ValidateScan = OK
End Function

Private Sub tvResult_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo errTrap

    Select Case Button
        Case vbRightButton:
                If Left(tvResult.SelectedItem.Key, 4) = "Port" Then
                    mnuThisPort.Visible = True
                Else
                    mnuThisPort.Visible = False
                End If
                
                PopupMenu mnuResult
    End Select

errTrap:
    On Error GoTo 0
    Exit Sub
End Sub

Private Sub txtManual_Validate(Cancel As Boolean)
    WritePrivateProfileString "PORT", "ManualString", txtManual.Text, App.Path & "\Conf.ini"
End Sub

Private Function CopyIco(ByVal FileName As String)
Dim FSO As FileSystemObject, IcoFile As File, IcoFoldR As Folder

    Set FSO = New FileSystemObject
    Set IcoFile = FSO.GetFile(FileName)
    Set IcoFoldR = FSO.GetFolder(App.Path & "\Ico")
    
        IcoFoldR.Copy IcoFile.ParentFolder
    
    Set FSO = Nothing
    Set IcoFile = Nothing
    Set IcoFoldR = Nothing
End Function

'RESOLVE DEFAULT SERVICE ON A PORT'
Private Function ResolvePort(ByVal Port As Long) As String
    Select Case Port
        Case 7:
                ResolvePort = "7/TCP [Echo]"
                Exit Function
        Case 11:
                ResolvePort = "11/TCP [Systat]"
                Exit Function
        Case 19:
                ResolvePort = "19/TCP [chargen]"
                Exit Function
        Case 21:
                ResolvePort = "21/TCP [FTP-Data]"
                Exit Function
        Case 22:
                ResolvePort = "22/TCP [SSH]"
                Exit Function
        Case 23:
                ResolvePort = "23/TCP [Telnet]"
                Exit Function
        Case 25:
                ResolvePort = "25/TCP [SMTP]"
                Exit Function
        Case 42:
                ResolvePort = "42/TCP [nameserver]"
                Exit Function
        Case 43:
                ResolvePort = "43/TCP [whois]"
                Exit Function
        Case 52:
                ResolvePort = "52/TCP [xns-time]"
                Exit Function
        Case 53:
                ResolvePort = "53/TCP [dns-zone]"
                Exit Function
        Case 63:
                ResolvePort = "63/TCP [whois++]"
                Exit Function
        Case 66:
                ResolvePort = "66/TCP [oracle-sqlnet]"
                Exit Function
        Case 70:
                ResolvePort = "70/TCP [gopher]"
                Exit Function
        Case 79:
                ResolvePort = "79/TCP [finger]"
                Exit Function
        Case 80:
                ResolvePort = "80/TCP [HTTP]"
                Exit Function
        Case 81:
                ResolvePort = "81/TCP [Alternate HTTP]"
                Exit Function
        Case 82:
                ResolvePort = "82/TCP [kerberos]"
                Exit Function
        Case 109:
                ResolvePort = "109/TCP [pop2]"
                Exit Function
        Case 110:
                ResolvePort = "110/TCP [pop3]"
                Exit Function
        Case 111:
                ResolvePort = "111/TCP [sunrpc]"
                Exit Function
        Case 118:
                ResolvePort = "118/TCP [sqlserv]"
                Exit Function
        Case 119:
                ResolvePort = "119/TCP [nntp]"
                Exit Function
        Case 123:
                ResolvePort = "123/TCP [ntp]"
                Exit Function
        Case 135:
                ResolvePort = "135/TCP [ntrpc-or-dce(epmap)]"
                Exit Function
        Case 137:
                ResolvePort = "137/TCP [netbios-ns]"
                Exit Function
        Case 138:
                ResolvePort = "138/TCP [netbios-dgm]"
                Exit Function
        Case 139:
                ResolvePort = "139/TCP [netbios]"
                Exit Function
        Case 143:
                ResolvePort = "143/TCP [imap]"
                Exit Function
        Case 177:
                ResolvePort = "177/TCP [xdmcp]"
                Exit Function
        Case 256:
                ResolvePort = "256/TCP [snmp-checkpoint]"
                Exit Function
        Case 389:
                ResolvePort = "389/TCP [ldap]"
                Exit Function
        Case 396:
                ResolvePort = "396/TCP [netware-ip]"
                Exit Function
        Case 443:
                ResolvePort = "443/TCP [HTTPS/SSL]"
                Exit Function
        Case 445:
                ResolvePort = "445/TCP [ms-smb-alternate]"
                Exit Function
        Case 512:
                ResolvePort = "512/TCP [exec]"
                Exit Function
        Case 513:
                ResolvePort = "513/TCP [rlogin]"
                Exit Function
        Case 514:
                ResolvePort = "514/TCP [rshell]"
                Exit Function
        Case 515:
                ResolvePort = "515/TCP [printer]"
                Exit Function
        Case 518:
                ResolvePort = "518/TCP [ntalk]"
                Exit Function
        Case 524:
                ResolvePort = "524/TCP [netware-ncp]"
                Exit Function
        Case 529:
                ResolvePort = "529/TCP [irc-serv]"
                Exit Function
        Case 901:
                ResolvePort = "901/TCP [samba-swat]"
                Exit Function
        Case 1024 To 1030:
                ResolvePort = Port & "/TCP [Echo]"
                Exit Function
        Case 1080:
                ResolvePort = "1080/TCP [socks]"
                Exit Function
        Case 1433:
                ResolvePort = "1433/TCP [ms-sql]"
                Exit Function
        Case 1498:
                ResolvePort = "1498/TCP [sybase-sql-anywhere]"
                Exit Function
        Case 1525:
                ResolvePort = "1525/TCP [oracle-srv]"
                Exit Function
        Case 1527:
                ResolvePort = "1527/TCP [oracle-tli]"
                Exit Function
        Case 1745:
                ResolvePort = "1745/TCP [winsock-proxy]"
                Exit Function
        Case 2001:
                ResolvePort = "2001/TCP [cisco-mgmt]"
                Exit Function
        Case 2049:
                ResolvePort = "2049/TCP [nfs]"
                Exit Function
        Case 2368:
                ResolvePort = "2368/TCP [sybase]"
                Exit Function
        Case 3001:
                ResolvePort = "3001/TCP [nessusd]"
                Exit Function
        Case 3306:
                ResolvePort = "3306/TCP [mysql]"
                Exit Function
        Case 3351:
                ResolvePort = "3351/TCP [ssql]"
                Exit Function
        Case 3389:
                ResolvePort = "3389/TCP [ms-termserv]"
                Exit Function
        Case 4001:
                ResolvePort = "4001/TCP [cisco-mgmt]"
                Exit Function
        Case 4045:
                ResolvePort = "4045/TCP [nfs-lockd]"
                Exit Function
        Case 4321:
                ResolvePort = "4321/TCP [rwhois]"
                Exit Function
        Case 5631:
                ResolvePort = "5631/TCP [pcanywhere]"
                Exit Function
        Case 5800:
                ResolvePort = "5800/TCP [vnc]"
                Exit Function
        Case 6000:
                ResolvePort = "6000/TCP [xwindows]"
                Exit Function
        Case 12345:
                ResolvePort = "12345/TCP [netbus]"
                Exit Function
        Case 32771:
                ResolvePort = "32771/TCP [rpc-solaris]"
                Exit Function
        Case 32780:
                ResolvePort = "32780/TCP [snmp-solaris]"
                Exit Function
        Case 54320:
                ResolvePort = "54320/TCP [bo2k]"
                Exit Function
        Case Else:
                ResolvePort = Port & "/TCP [?????????]"
                Exit Function
    End Select
End Function


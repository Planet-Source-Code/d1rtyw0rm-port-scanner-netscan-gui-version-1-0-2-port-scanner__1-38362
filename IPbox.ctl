VERSION 5.00
Begin VB.UserControl IPbox 
   BackColor       =   &H8000000D&
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2565
   ScaleHeight     =   735
   ScaleWidth      =   2565
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   4
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   960
         MaxLength       =   3
         TabIndex        =   3
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   480
         MaxLength       =   3
         TabIndex        =   2
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   0
         MaxLength       =   3
         TabIndex        =   1
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "."
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "."
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   0
         Width           =   135
      End
      Begin VB.Label lblPoint 
         BackColor       =   &H80000005&
         Caption         =   "."
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   0
         Width           =   135
      End
   End
End
Attribute VB_Name = "IPbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Event Declarations:
'Event Click()
'Event DblClick()
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
'Déclarations d'événements:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
'Valeurs de propriétés par défaut:
'Const m_def_FillColor = vbBlack
'Const m_def_Byte0 = ""
'Const m_def_Byte2 = ""
'Const m_def_Byte1 = ""
'Const m_def_Byte3 = ""
'Const m_def_FillColor = vbBlack
'Const m_def_Text = ""
'Variables de propriétés:
'Dim m_FillColor As OLE_COLOR
'Dim m_Byte0 As String
'Dim m_Byte2 As String
'Dim m_Byte1 As String
'Dim m_Byte3 As String
'Dim m_FillColor As OLE_COLOR
'Dim m_Text As String





Private Sub txtIP_KeyPress(Index As Integer, KeyAscii As Integer)
    ' N'ACCEPTE QUE LES NOMBRES '
    If InStr("0123456789" & Chr(8) & Chr(32), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    'N'ACCEPTE PAS LES VALEUR SUPÉRIEUR A 255'
    If Val(txtIP(Index).Text & Chr(KeyAscii)) > 255 Then
        KeyAscii = 0
    End If
    
    'SWITCH FOCUS AFTER 3 CHAR OR SPACE'
    If Index < 3 Then
        If KeyAscii = 32 Then txtIP(Index + 1).SetFocus
        
        If Len(txtIP(Index)) > 1 Then
            txtIP(Index + 1).SetFocus
        End If
    End If
    
End Sub

'RETURN VAL'
Public Property Get AdresseIP() As String
    AdresseIP = txtIP(0).Text & "." & txtIP(1).Text & "." & txtIP(2).Text & "." & txtIP(3).Text
End Property

'PROPERTY COULEUR'
Public Property Let CouleurFond(ByVal nCouleur As OLE_COLOR)
    UserControl.BackColor = nCouleur
End Property

Public Property Get CouleurFond() As OLE_COLOR
    CouleurFond = UserControl.BackColor
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    PropBag.WriteProperty "CouleurFond", UserControl.BackColor

    Call PropBag.WriteProperty("BorderStyle", Picture1.BorderStyle, 1)
    Call PropBag.WriteProperty("Appearance", Picture1.Appearance, 1)
    Call PropBag.WriteProperty("Enabled", Picture1.Enabled, True)
    Call PropBag.WriteProperty("FillColor", txtIP(0).ForeColor, &H80000008)
    Call PropBag.WriteProperty("BackColor", Picture1.BackColor, &H8000000D)
    Call PropBag.WriteProperty("Locked", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Text", txtIP(0).Text, "")
'    Call PropBag.WriteProperty("FillColor", m_FillColor, m_def_FillColor)
'    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("Byte0", txtIP(0).Text, "")
    Call PropBag.WriteProperty("Byte2", txtIP(2).Text, "")
    Call PropBag.WriteProperty("Byte1", txtIP(1).Text, "")
    Call PropBag.WriteProperty("Byte3", txtIP(3).Text, "")
'    Call PropBag.WriteProperty("FillColor", m_FillColor, m_def_FillColor)
'    Call PropBag.WriteProperty("Byte0", m_Byte0, m_def_Byte0)
'    Call PropBag.WriteProperty("Byte2", m_Byte2, m_def_Byte2)
'    Call PropBag.WriteProperty("Byte1", m_Byte1, m_def_Byte1)
'    Call PropBag.WriteProperty("Byte3", m_Byte3, m_def_Byte3)
    Call PropBag.WriteProperty("Byte0_FillColor", txtIP(0).ForeColor, &H80000008)
    Call PropBag.WriteProperty("Byte1_FillColor", txtIP(1).ForeColor, &H80000008)
    Call PropBag.WriteProperty("Byte2_FillColor", txtIP(2).ForeColor, &H80000008)
    Call PropBag.WriteProperty("Byte3_FillColor", txtIP(3).ForeColor, &H80000008)
    Call PropBag.WriteProperty("Byte0", txtIP(0).Text, "")
    Call PropBag.WriteProperty("Byte2", txtIP(2).Text, "")
    Call PropBag.WriteProperty("Byte1", txtIP(1).Text, "")
    Call PropBag.WriteProperty("Byte3", txtIP(3).Text, "")
    Call PropBag.WriteProperty("Byte0_FillColor", txtIP(0).ForeColor, &H80000008)
    Call PropBag.WriteProperty("Byte1_FillColor", txtIP(1).ForeColor, &H80000008)
    Call PropBag.WriteProperty("Byte2_FillColor", txtIP(2).ForeColor, &H80000008)
    Call PropBag.WriteProperty("Byte3_FillColor", txtIP(3).ForeColor, &H80000008)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("CouleurFond", vbButtonFace)

    Picture1.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    Picture1.Appearance = PropBag.ReadProperty("Appearance", 1)
    Picture1.Enabled = PropBag.ReadProperty("Enabled", True)
    txtIP(0).ForeColor = PropBag.ReadProperty("FillColor", &H80000008)
    Picture1.BackColor = PropBag.ReadProperty("BackColor", &H8000000D)
    UserControl.Enabled = PropBag.ReadProperty("Locked", True)
    txtIP(0).Text = PropBag.ReadProperty("Text", "")
'    m_FillColor = PropBag.ReadProperty("FillColor", m_def_FillColor)
'    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    txtIP(0).Text = PropBag.ReadProperty("Byte0", "")
    txtIP(2).Text = PropBag.ReadProperty("Byte2", "")
    txtIP(1).Text = PropBag.ReadProperty("Byte1", "")
    txtIP(3).Text = PropBag.ReadProperty("Byte3", "")
'    m_FillColor = PropBag.ReadProperty("FillColor", m_def_FillColor)
'    m_Byte0 = PropBag.ReadProperty("Byte0", m_def_Byte0)
'    m_Byte2 = PropBag.ReadProperty("Byte2", m_def_Byte2)
'    m_Byte1 = PropBag.ReadProperty("Byte1", m_def_Byte1)
'    m_Byte3 = PropBag.ReadProperty("Byte3", m_def_Byte3)
    txtIP(0).ForeColor = PropBag.ReadProperty("Byte0_FillColor", &H80000008)
    txtIP(1).ForeColor = PropBag.ReadProperty("Byte1_FillColor", &H80000008)
    txtIP(2).ForeColor = PropBag.ReadProperty("Byte2_FillColor", &H80000008)
    txtIP(3).ForeColor = PropBag.ReadProperty("Byte3_FillColor", &H80000008)
    txtIP(0).Text = PropBag.ReadProperty("Byte0", "")
    txtIP(2).Text = PropBag.ReadProperty("Byte2", "")
    txtIP(1).Text = PropBag.ReadProperty("Byte1", "")
    txtIP(3).Text = PropBag.ReadProperty("Byte3", "")
    txtIP(0).ForeColor = PropBag.ReadProperty("Byte0_FillColor", &H80000008)
    txtIP(1).ForeColor = PropBag.ReadProperty("Byte1_FillColor", &H80000008)
    txtIP(2).ForeColor = PropBag.ReadProperty("Byte2_FillColor", &H80000008)
    txtIP(3).ForeColor = PropBag.ReadProperty("Byte3_FillColor", &H80000008)
End Sub

Private Sub UserControl_InitProperties()

    UserControl.BackColor = vbButtonFace

'    m_FillColor = m_def_FillColor
'    m_Text = m_def_Text
'    m_FillColor = m_def_FillColor
'    m_Byte0 = m_def_Byte0
'    m_Byte2 = m_def_Byte2
'    m_Byte1 = m_def_Byte1
'    m_Byte3 = m_def_Byte3
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = Picture1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    Picture1.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = Picture1.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    Picture1.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=Picture1,Picture1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Renvoie ou définit une valeur qui détermine si un objet peut répondre à des événements générés par l'utilisateur."
    Enabled = Picture1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Picture1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
'
''ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
''MappingInfo=txtIP(0),txtIP,0,ForeColor
'Public Property Get FillColor() As OLE_COLOR
'    FillColor = txtIP(0).ForeColor
'End Property
'
'Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
'    txtIP(0).ForeColor() = New_FillColor
'    PropertyChanged "FillColor"
'End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=Picture1,Picture1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Renvoie ou définit la couleur d'arrière-plan utilisée pour afficher le texte et les graphiques d'un objet."
    BackColor = Picture1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Picture1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Renvoie ou définit une valeur qui détermine si un objet peut répondre à des événements générés par l'utilisateur."
    Locked = UserControl.Enabled
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    UserControl.Enabled() = New_Locked
    PropertyChanged "Locked"
End Property
''
'''ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'''MappingInfo=txtIP(0),txtIP,0,Text
''Public Property Get Text() As String
''    Text = txtIP(0).Text
''End Property
''
''Public Property Let Text(ByVal New_Text As String)
''    txtIP(0).Text() = New_Text
''    PropertyChanged "Text"
''End Property
''
''ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
''MemberInfo=10,0,0,
'Public Property Get FillColor() As OLE_COLOR
'    FillColor = m_FillColor
'End Property
'
'Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
'    m_FillColor = New_FillColor
'    PropertyChanged "FillColor"
'End Property
'
''ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
''MemberInfo=13,0,0,
'Public Property Get Text() As String
'    Text = m_Text
'End Property
'
'Public Property Let Text(ByVal New_Text As String)
'    m_Text = New_Text
'    PropertyChanged "Text"
'End Property
''
'''ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'''MappingInfo=txtIP(0),txtIP,0,Text
''Public Property Get Byte0() As String
''    Byte0 = txtIP(0).Text
''End Property
''
''Public Property Let Byte0(ByVal New_Byte0 As String)
''    txtIP(0).Text() = New_Byte0
''    PropertyChanged "Byte0"
''End Property
''
'''ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'''MappingInfo=txtIP(2),txtIP,2,Text
''Public Property Get Byte2() As String
''    Byte2 = txtIP(2).Text
''End Property
''
''Public Property Let Byte2(ByVal New_Byte2 As String)
''    txtIP(2).Text() = New_Byte2
''    PropertyChanged "Byte2"
''End Property
''
'''ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'''MappingInfo=txtIP(1),txtIP,1,Text
''Public Property Get Byte1() As String
''    Byte1 = txtIP(1).Text
''End Property
''
''Public Property Let Byte1(ByVal New_Byte1 As String)
''    txtIP(1).Text() = New_Byte1
''    PropertyChanged "Byte1"
''End Property
''
'''ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'''MappingInfo=txtIP(3),txtIP,3,Text
''Public Property Get Byte3() As String
''    Byte3 = txtIP(3).Text
''End Property
''
''Public Property Let Byte3(ByVal New_Byte3 As String)
''    txtIP(3).Text() = New_Byte3
''    PropertyChanged "Byte3"
''End Property
''
''ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
''MemberInfo=10,0,0,
'Public Property Get FillColor() As OLE_COLOR
'    FillColor = m_FillColor
'End Property
'
'Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
'    m_FillColor = New_FillColor
'    PropertyChanged "FillColor"
'End Property
'
''ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
''MemberInfo=13,0,0,
'Public Property Get Byte0() As String
'    Byte0 = m_Byte0
'End Property
'
'Public Property Let Byte0(ByVal New_Byte0 As String)
'    m_Byte0 = New_Byte0
'    PropertyChanged "Byte0"
'End Property
'
''ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
''MemberInfo=13,0,0,
'Public Property Get Byte2() As String
'    Byte2 = m_Byte2
'End Property
'
'Public Property Let Byte2(ByVal New_Byte2 As String)
'    m_Byte2 = New_Byte2
'    PropertyChanged "Byte2"
'End Property
'
''ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
''MemberInfo=13,0,0,
'Public Property Get Byte1() As String
'    Byte1 = m_Byte1
'End Property
'
'Public Property Let Byte1(ByVal New_Byte1 As String)
'    m_Byte1 = New_Byte1
'    PropertyChanged "Byte1"
'End Property
'
''ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
''MemberInfo=13,0,0,
'Public Property Get Byte3() As String
'    Byte3 = m_Byte3
'End Property
'
'Public Property Let Byte3(ByVal New_Byte3 As String)
'    m_Byte3 = New_Byte3
'    PropertyChanged "Byte3"
'End Property
'
''ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
''MappingInfo=txtIP(0),txtIP,0,ForeColor
'Public Property Get Byte0_FillColor() As OLE_COLOR
'    Byte0_FillColor = txtIP(0).ForeColor
'End Property
'
'Public Property Let Byte0_FillColor(ByVal New_Byte0_FillColor As OLE_COLOR)
'    txtIP(0).ForeColor() = New_Byte0_FillColor
'    PropertyChanged "Byte0_FillColor"
'End Property
'
''ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
''MappingInfo=txtIP(1),txtIP,1,ForeColor
'Public Property Get Byte1_FillColor() As OLE_COLOR
'    Byte1_FillColor = txtIP(1).ForeColor
'End Property
'
'Public Property Let Byte1_FillColor(ByVal New_Byte1_FillColor As OLE_COLOR)
'    txtIP(1).ForeColor() = New_Byte1_FillColor
'    PropertyChanged "Byte1_FillColor"
'End Property
'
''ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
''MappingInfo=txtIP(2),txtIP,2,ForeColor
'Public Property Get Byte2_FillColor() As OLE_COLOR
'    Byte2_FillColor = txtIP(2).ForeColor
'End Property
'
'Public Property Let Byte2_FillColor(ByVal New_Byte2_FillColor As OLE_COLOR)
'    txtIP(2).ForeColor() = New_Byte2_FillColor
'    PropertyChanged "Byte2_FillColor"
'End Property
'
''ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
''MappingInfo=txtIP(3),txtIP,3,ForeColor
'Public Property Get Byte3_FillColor() As OLE_COLOR
'    Byte3_FillColor = txtIP(3).ForeColor
'End Property
'
'Public Property Let Byte3_FillColor(ByVal New_Byte3_FillColor As OLE_COLOR)
'    txtIP(3).ForeColor() = New_Byte3_FillColor
'    PropertyChanged "Byte3_FillColor"
'End Property
'
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=txtIP(0),txtIP,0,Text
Public Property Get Byte0() As String
Attribute Byte0.VB_Description = "Renvoie ou définit le texte contenu dans le contrôle."
    Byte0 = txtIP(0).Text
End Property

Public Property Let Byte0(ByVal New_Byte0 As String)
    txtIP(0).Text() = New_Byte0
    PropertyChanged "Byte0"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=txtIP(2),txtIP,2,Text
Public Property Get Byte2() As String
Attribute Byte2.VB_Description = "Renvoie ou définit le texte contenu dans le contrôle."
    Byte2 = txtIP(2).Text
End Property

Public Property Let Byte2(ByVal New_Byte2 As String)
    txtIP(2).Text() = New_Byte2
    PropertyChanged "Byte2"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=txtIP(1),txtIP,1,Text
Public Property Get Byte1() As String
Attribute Byte1.VB_Description = "Renvoie ou définit le texte contenu dans le contrôle."
    Byte1 = txtIP(1).Text
End Property

Public Property Let Byte1(ByVal New_Byte1 As String)
    txtIP(1).Text() = New_Byte1
    PropertyChanged "Byte1"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=txtIP(3),txtIP,3,Text
Public Property Get Byte3() As String
Attribute Byte3.VB_Description = "Renvoie ou définit le texte contenu dans le contrôle."
    Byte3 = txtIP(3).Text
End Property

Public Property Let Byte3(ByVal New_Byte3 As String)
    txtIP(3).Text() = New_Byte3
    PropertyChanged "Byte3"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=txtIP(0),txtIP,0,ForeColor
Public Property Get Byte0_FillColor() As OLE_COLOR
Attribute Byte0_FillColor.VB_Description = "Renvoie ou définit la couleur de premier plan utilisée pour afficher le texte et les graphiques d'un objet."
    Byte0_FillColor = txtIP(0).ForeColor
End Property

Public Property Let Byte0_FillColor(ByVal New_Byte0_FillColor As OLE_COLOR)
    txtIP(0).ForeColor() = New_Byte0_FillColor
    PropertyChanged "Byte0_FillColor"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=txtIP(1),txtIP,1,ForeColor
Public Property Get Byte1_FillColor() As OLE_COLOR
Attribute Byte1_FillColor.VB_Description = "Renvoie ou définit la couleur de premier plan utilisée pour afficher le texte et les graphiques d'un objet."
    Byte1_FillColor = txtIP(1).ForeColor
End Property

Public Property Let Byte1_FillColor(ByVal New_Byte1_FillColor As OLE_COLOR)
    txtIP(1).ForeColor() = New_Byte1_FillColor
    PropertyChanged "Byte1_FillColor"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=txtIP(2),txtIP,2,ForeColor
Public Property Get Byte2_FillColor() As OLE_COLOR
Attribute Byte2_FillColor.VB_Description = "Renvoie ou définit la couleur de premier plan utilisée pour afficher le texte et les graphiques d'un objet."
    Byte2_FillColor = txtIP(2).ForeColor
End Property

Public Property Let Byte2_FillColor(ByVal New_Byte2_FillColor As OLE_COLOR)
    txtIP(2).ForeColor() = New_Byte2_FillColor
    PropertyChanged "Byte2_FillColor"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=txtIP(3),txtIP,3,ForeColor
Public Property Get Byte3_FillColor() As OLE_COLOR
Attribute Byte3_FillColor.VB_Description = "Renvoie ou définit la couleur de premier plan utilisée pour afficher le texte et les graphiques d'un objet."
    Byte3_FillColor = txtIP(3).ForeColor
End Property

Public Property Let Byte3_FillColor(ByVal New_Byte3_FillColor As OLE_COLOR)
    txtIP(3).ForeColor() = New_Byte3_FillColor
    PropertyChanged "Byte3_FillColor"
End Property


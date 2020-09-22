VERSION 5.00
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "dhtmled.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl xHTML 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   3945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   ScaleHeight     =   3945
   ScaleWidth      =   11250
   ToolboxBitmap   =   "xHMTL.ctx":0000
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   11250
      _CBHeight       =   390
      _Version        =   "6.0.8169"
      Child1          =   "Toolbar2"
      MinHeight1      =   330
      Width1          =   6120
      NewRow1         =   0   'False
      Child2          =   "pctFonts"
      MinHeight2      =   330
      Width2          =   4740
      NewRow2         =   0   'False
      Begin VB.PictureBox pctFonts 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   6315
         ScaleHeight     =   330
         ScaleWidth      =   4845
         TabIndex        =   3
         Top             =   30
         Width           =   4845
         Begin VB.ComboBox cmbParagraph 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "xHMTL.ctx":0312
            Left            =   2985
            List            =   "xHMTL.ctx":0331
            Style           =   2  'Dropdown List
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   15
            Width           =   1890
         End
         Begin VB.ComboBox cmbFontSize 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "xHMTL.ctx":03BE
            Left            =   1830
            List            =   "xHMTL.ctx":03DA
            Style           =   2  'Dropdown List
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   15
            Width           =   1170
         End
         Begin VB.ComboBox cmbFont 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "xHMTL.ctx":042B
            Left            =   0
            List            =   "xHMTL.ctx":043E
            Style           =   2  'Dropdown List
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   15
            Width           =   1830
         End
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ToolBar1Images"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   18
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bold"
               Description     =   "To change the selected text in Bold....."
               Object.ToolTipText     =   "Bold"
               ImageKey        =   "b"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Italic"
               Description     =   "To change the selected text in Italic....."
               Object.ToolTipText     =   "Italic"
               ImageKey        =   "i"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Underline"
               Description     =   "To change the selected text in Underline....."
               Object.ToolTipText     =   "Underline"
               ImageKey        =   "u"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "undo"
               Object.ToolTipText     =   "Undo Last Item"
               ImageKey        =   "Undo"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "redo"
               Object.ToolTipText     =   "Redo Last undone"
               ImageKey        =   "Redo"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Align Left"
               Description     =   "To Align the text in left...."
               Object.ToolTipText     =   "Align Left"
               ImageKey        =   "Align Left"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Center"
               Description     =   "To Align the text in Center...."
               Object.ToolTipText     =   "Center"
               ImageKey        =   "Center"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Align Right"
               Description     =   "To Align the text in right...."
               Object.ToolTipText     =   "Align Right"
               ImageKey        =   "Align Right"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Indent"
               Description     =   "To indent the selected text...."
               Object.ToolTipText     =   "Indent"
               ImageKey        =   "indent"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Outdent"
               Description     =   "To outdent the selected text...."
               Object.ToolTipText     =   "Outdent"
               ImageKey        =   "outdent"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "numbull"
               ImageKey        =   "numbull"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "dotbull"
               ImageKey        =   "dotbull"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "hr"
               ImageKey        =   "rule"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "forecolor"
               Object.ToolTipText     =   "Forecolor"
               ImageKey        =   "forecolor"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "backcolor"
               Object.ToolTipText     =   "backcolor"
               ImageKey        =   "backcolor"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Object.ToolTipText     =   "Save As Template"
               ImageKey        =   "save"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               Object.ToolTipText     =   "Open Existing Template"
               ImageKey        =   "open"
            EndProperty
         EndProperty
      End
   End
   Begin DHTMLEDLibCtl.DHTMLEdit DHTMLSafe1 
      Height          =   450
      Left            =   1335
      TabIndex        =   0
      Top             =   1665
      Width           =   480
      ActivateApplets =   0   'False
      ActivateActiveXControls=   0   'False
      ActivateDTCs    =   -1  'True
      ShowDetails     =   0   'False
      ShowBorders     =   0   'False
      Appearance      =   1
      Scrollbars      =   -1  'True
      ScrollbarAppearance=   1
      SourceCodePreservation=   -1  'True
      AbsoluteDropMode=   0   'False
      SnapToGrid      =   0   'False
      SnapToGridX     =   50
      SnapToGridY     =   50
      BrowseMode      =   -1  'True
      UseDivOnCarriageReturn=   0   'False
   End
   Begin MSComctlLib.ImageList ToolBar1Images 
      Left            =   4350
      Top             =   435
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":0485
            Key             =   "u"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":0597
            Key             =   "i"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":06A9
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":07BB
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":08CD
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":09DF
            Key             =   "b"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":0AF1
            Key             =   "Copy"
            Object.Tag             =   "&Copy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":1033
            Key             =   "Cut"
            Object.Tag             =   "Cu&t"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":1575
            Key             =   "Paste"
            Object.Tag             =   "&Paste"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":1AB7
            Key             =   "Redo"
            Object.Tag             =   "&Redo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":1BC9
            Key             =   "Undo"
            Object.Tag             =   "&Undo"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":1CDB
            Key             =   "numbull"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":202D
            Key             =   "dotbull"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":237F
            Key             =   "pic"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":31D1
            Key             =   "backcolor"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":3523
            Key             =   "rule"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":3875
            Key             =   "forecolor"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":46C7
            Key             =   "indent"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":4A19
            Key             =   "outdent"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":4D6B
            Key             =   "save"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xHMTL.ctx":52AD
            Key             =   "open"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CDAttach 
      Left            =   5010
      Top             =   510
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.*"
      DialogTitle     =   "Attachment Filename"
      FileName        =   "*.*"
      Filter          =   "*.*"
   End
End
Attribute VB_Name = "xHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event openClicked()
Event saveClicked()
Event ContextMenuAction(ByVal itemIndex As Long) 'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,ContextMenuAction
Event DisplayChanged() 'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,DisplayChanged
Event DocumentComplete() 'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,DocumentComplete
Event onblur() 'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,onblur
Event onclick() 'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,onclick
Event ondblclick() 'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,ondblclick
Event onkeydown() 'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,onkeydown
Event onkeypress() 'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,onkeypress
Event onkeyup() 'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,onkeyup
Event onmousedown() 'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,onmousedown
Event onmousemove() 'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,onmousemove
Event onmouseout() 'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,onmouseout
Event onmouseover() 'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,onmouseover
Event onmouseup() 'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,onmouseup
Event onreadystatechange() 'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,onreadystatechange
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event ShowContextMenu(ByVal xPos As Long, ByVal yPos As Long) 'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,ShowContextMenu
Event Validate(Cancel As Boolean) 'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,Validate
Attribute Validate.VB_Description = "Occurs when a control loses focus to a control that causes validation."
'Default Property Values:
Const m_def_ShowToolbar = True
'Property Variables:
Dim m_ShowToolbar As Boolean
Dim m_BrowseMode As Boolean
'
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "method Refresh"
    DHTMLSafe1.Refresh
End Sub


Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
    Call UserControl_Resize
End Sub

Private Sub DHTMLSafe1_DragDrop(Source As Control, X As Single, Y As Single)
    PropertyChanged "DocumentHTML"
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Ambient.UserMode = False Or DHTMLSafe1.BrowseMode = True Then
        Exit Sub
    End If
    Select Case (LCase(Button.Key))
        Case Is = "bold"
            DHTMLSafe1.ExecCommand DECMD_BOLD, OLECMDEXECOPT_DODEFAULT, Null
        Case Is = "italic"
            DHTMLSafe1.ExecCommand DECMD_ITALIC, OLECMDEXECOPT_DODEFAULT, Null
        Case Is = "undo"
            DHTMLSafe1.ExecCommand DECMD_UNDO, OLECMDEXECOPT_DODEFAULT, Null
        Case Is = "redo"
            DHTMLSafe1.ExecCommand DECMD_REDO, OLECMDEXECOPT_DODEFAULT, Null
        Case Is = "underline"
            DHTMLSafe1.ExecCommand DECMD_UNDERLINE, OLECMDEXECOPT_DODEFAULT, Null
        Case Is = "align right"
            DHTMLSafe1.ExecCommand DECMD_JUSTIFYRIGHT, OLECMDEXECOPT_DODEFAULT, Null
        Case Is = "align left"
            DHTMLSafe1.ExecCommand DECMD_JUSTIFYLEFT, OLECMDEXECOPT_DODEFAULT, Null
        Case Is = "center"
            DHTMLSafe1.ExecCommand DECMD_JUSTIFYCENTER, OLECMDEXECOPT_DODEFAULT, Null
        Case Is = "outdent"
            DHTMLSafe1.ExecCommand DECMD_OUTDENT, OLECMDEXECOPT_DODEFAULT, Null
        Case Is = "indent"
            DHTMLSafe1.ExecCommand DECMD_INDENT, OLECMDEXECOPT_DODEFAULT, Null
        Case Is = "pic"
            DHTMLSafe1.ExecCommand DECMD_IMAGE, OLECMDEXECOPT_PROMPTUSER, Null
        Case Is = "numbull"
            DHTMLSafe1.ExecCommand DECMD_ORDERLIST, OLECMDEXECOPT_DODEFAULT, Null
        Case Is = "dotbull"
            DHTMLSafe1.ExecCommand DECMD_UNORDERLIST, OLECMDEXECOPT_DODEFAULT, Null
        Case Is = "hr"
            'DHTMLSafe1.ExecCommand DECMD_, OLECMDEXECOPT_DODEFAULT, Null
            DHTMLSafe1.DOM.ExecCommand "inserthorizontalrule", False, Null
        Case Is = "open"
            RaiseEvent openClicked
        Case Is = "save"
            RaiseEvent saveClicked
        Case Is = "forecolor", "backcolor"
            On Local Error GoTo noColorSeleded
            CDAttach.ShowColor
            DHTMLSafe1.DOM.ExecCommand LCase(Button.Key), False, CDAttach.Color
    End Select
noColorSeleded:
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub
'

Private Sub cmbParagraph_Click()
On Error Resume Next
    If DHTMLSafe1.BrowseMode = True Then
        Exit Sub
    End If
    Dim paraStyle As String
    If cmbParagraph.ListIndex = 0 Then
        paraStyle = "Normal"
    ElseIf cmbParagraph.ListIndex > 0 And cmbParagraph.ListIndex < 8 Then
        paraStyle = "Heading " & cmbParagraph.ListIndex
    ElseIf cmbParagraph.ListIndex = 8 Then
        paraStyle = "Address"
    ElseIf cmbParagraph.ListIndex = 9 Then
        paraStyle = "Formatted"
    End If
    DHTMLSafe1.DOM.ExecCommand "formatBlock", False, paraStyle
    
End Sub

Private Sub cmbFont_Click()
On Local Error Resume Next
    If DHTMLSafe1.BrowseMode = True Then
        Exit Sub
    End If
    DHTMLSafe1.DOM.ExecCommand "fontname", False, cmbFont.Text
End Sub
Private Sub cmbFontSize_Click()
On Local Error Resume Next
    If DHTMLSafe1.BrowseMode = True Then
        Exit Sub
    End If
    DHTMLSafe1.DOM.ExecCommand "fontsize", False, cmbFontSize.ItemData(cmbFontSize.ListIndex)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,AbsoluteDropMode
Public Property Get AbsoluteDropMode() As Boolean
Attribute AbsoluteDropMode.VB_Description = "property AbsoluteDropMode"
    AbsoluteDropMode = DHTMLSafe1.AbsoluteDropMode
End Property

Public Property Let AbsoluteDropMode(ByVal New_AbsoluteDropMode As Boolean)
    DHTMLSafe1.AbsoluteDropMode() = New_AbsoluteDropMode
    PropertyChanged "AbsoluteDropMode"
End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,BaseURL
Public Property Get BaseURL() As String
Attribute BaseURL.VB_Description = "property BaseURL"
    BaseURL = DHTMLSafe1.BaseURL
End Property

Public Property Let BaseURL(ByVal New_BaseURL As String)
    DHTMLSafe1.BaseURL() = New_BaseURL
    PropertyChanged "BaseURL"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,Busy
Public Property Get Busy() As Boolean
Attribute Busy.VB_Description = "property Busy"
    Busy = DHTMLSafe1.Busy
End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function CanPropertyChange(ByVal PropertyName As String) As Boolean
Attribute CanPropertyChange.VB_Description = "Asks the container if a property bound to a data source can be changed.  The CanPropertyChange method is most useful if the property specified in PropertyName is bound to a data source."
    CanPropertyChange = True
End Function
'
Private Sub DHTMLSafe1_ContextMenuAction(ByVal itemIndex As Long)
    RaiseEvent ContextMenuAction(itemIndex)
End Sub
'

Private Sub DHTMLSafe1_DisplayChanged()
    RaiseEvent DisplayChanged
End Sub

Private Sub DHTMLSafe1_DocumentComplete()
    RaiseEvent DocumentComplete
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,DocumentHTML
Public Property Get DocumentHTML() As String
Attribute DocumentHTML.VB_Description = "property DocumentHTML"
Attribute DocumentHTML.VB_MemberFlags = "163c"
    DocumentHTML = DHTMLSafe1.DocumentHTML
End Property

Public Property Let DocumentHTML(ByVal New_DocumentHTML As String)
    DHTMLSafe1.DocumentHTML() = New_DocumentHTML
    PropertyChanged "DocumentHTML"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,DocumentTitle
Public Property Get DocumentTitle() As String
Attribute DocumentTitle.VB_Description = "property DocumentTitle"
    DocumentTitle = DHTMLSafe1.DocumentTitle
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,ExecCommand
Public Function ExecCommand(cmdID As DHTMLEDITCMDID, cmdexecopt As OLECMDEXECOPT, Optional pInVar As Variant) As Variant
Attribute ExecCommand.VB_Description = "method ExecCommand"
    ExecCommand = DHTMLSafe1.ExecCommand(cmdID, cmdexecopt, pInVar)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,FilterSourceCode
Public Function FilterSourceCode(ByVal sourceCodeIn As String) As String
Attribute FilterSourceCode.VB_Description = "method FilterSourceCode"
    FilterSourceCode = DHTMLSafe1.FilterSourceCode(sourceCodeIn)
End Function
'

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,IsDirty
Public Property Get IsDirty() As Boolean
Attribute IsDirty.VB_Description = "property IsDirty"
    IsDirty = DHTMLSafe1.IsDirty
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,LoadURL
Public Sub LoadURL(ByVal url As String)
Attribute LoadURL.VB_Description = "method LoadURL"
    DHTMLSafe1.LoadURL url
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,NewDocument
Public Sub NewDocument()
Attribute NewDocument.VB_Description = "method NewDocument"
    DHTMLSafe1.NewDocument
End Sub
'

Private Sub DHTMLSafe1_onblur()
    If Ambient.UserMode = True Then
        RaiseEvent onblur
    End If
End Sub

Private Sub DHTMLSafe1_onclick()
    If Ambient.UserMode = True Then
        RaiseEvent onclick
    End If
End Sub

Private Sub DHTMLSafe1_ondblclick()
    If Ambient.UserMode = True Then
        RaiseEvent ondblclick
    End If
End Sub

Private Sub DHTMLSafe1_onkeydown()
    If Ambient.UserMode = True Then
        PropertyChanged "DocumentHTML"
        RaiseEvent onkeydown
    End If
End Sub

Private Sub DHTMLSafe1_onkeypress()
    If Ambient.UserMode = True Then
        PropertyChanged "DocumentHTML"
        RaiseEvent onkeypress
    End If
End Sub

Private Sub DHTMLSafe1_onkeyup()
    If Ambient.UserMode = True Then
        RaiseEvent onkeyup
    End If
End Sub

Private Sub DHTMLSafe1_onmousedown()
    If Ambient.UserMode = True Then
        RaiseEvent onmousedown
    End If
End Sub

Private Sub DHTMLSafe1_onmousemove()
    If Ambient.UserMode = True Then
        RaiseEvent onmousemove
    End If
End Sub

Private Sub DHTMLSafe1_onmouseout()
    If Ambient.UserMode = True Then
        RaiseEvent onmouseout
    End If
End Sub

Private Sub DHTMLSafe1_onmouseover()
    If Ambient.UserMode = True Then
        RaiseEvent onmouseover
    End If
End Sub

Private Sub DHTMLSafe1_onmouseup()
    If Ambient.UserMode = True Then
        RaiseEvent onmouseup
    End If
End Sub

Private Sub DHTMLSafe1_onreadystatechange()
    If Ambient.UserMode = True Then
        RaiseEvent onreadystatechange
    End If
End Sub
'

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,QueryStatus
Public Function QueryStatus(cmdID As DHTMLEDITCMDID) As DHTMLEDITCMDF
Attribute QueryStatus.VB_Description = "method QueryStatus"
    If Ambient.UserMode = True Then
        QueryStatus = DHTMLSafe1.QueryStatus(cmdID)
    End If
End Function

Private Sub UserControl_GotFocus()
    If Ambient.UserMode = True Then
        DHTMLSafe1.SetFocus
    End If
End Sub

Private Sub UserControl_InitProperties()
    DHTMLSafe1.ActivateDTCs = False
    showPrintValue "UserControl_InitProperties()"
    m_BrowseMode = False
    m_ShowToolbar = m_def_ShowToolbar
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    If Ambient.UserMode = False Then
        DHTMLSafe1.BrowseMode = True
        DisbaleToolbar True
    Else
        DHTMLSafe1.BrowseMode = Me.BrowseMode
    End If
    If Me.ShowToolbar = True Then
        DHTMLSafe1.Move 0, CoolBar1.Height + 20, UserControl.ScaleWidth, UserControl.ScaleHeight - (CoolBar1.Height + 20)
    Else
        DHTMLSafe1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    End If
    RaiseEvent Resize
    showPrintValue "UserControl_Resize()"
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,ScaleHeight
'Public Property Get ScaleHeight() As Single
'    ScaleHeight = UserControl.ScaleHeight
'End Property
'
'Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
'    UserControl.ScaleHeight() = New_ScaleHeight
'    PropertyChanged "ScaleHeight"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleLeft
Public Property Get ScaleLeft() As Single
Attribute ScaleLeft.VB_Description = "Returns/sets the horizontal coordinates for the left edges of an object."
    ScaleLeft = UserControl.ScaleLeft
End Property

Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
    UserControl.ScaleLeft() = New_ScaleLeft
    PropertyChanged "ScaleLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleMode
Public Property Get ScaleMode() As Integer
Attribute ScaleMode.VB_Description = "Returns/sets a value indicating measurement units for object coordinates when using graphics methods or positioning controls."
    ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)
    UserControl.ScaleMode() = New_ScaleMode
    PropertyChanged "ScaleMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleTop
Public Property Get ScaleTop() As Single
Attribute ScaleTop.VB_Description = "Returns/sets the vertical coordinates for the top edges of an object."
    ScaleTop = UserControl.ScaleTop
End Property

Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
    UserControl.ScaleTop() = New_ScaleTop
    PropertyChanged "ScaleTop"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    UserControl.ScaleWidth() = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,ScrollbarAppearance
Public Property Get ScrollbarAppearance() As DHTMLEDITAPPEARANCE
Attribute ScrollbarAppearance.VB_Description = "property ScrollbarAppearance"
    ScrollbarAppearance = DHTMLSafe1.ScrollbarAppearance
End Property

Public Property Let ScrollbarAppearance(ByVal New_ScrollbarAppearance As DHTMLEDITAPPEARANCE)
    DHTMLSafe1.ScrollbarAppearance() = New_ScrollbarAppearance
    PropertyChanged "ScrollbarAppearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,Scrollbars
Public Property Get Scrollbars() As Boolean
Attribute Scrollbars.VB_Description = "property Scrollbars"
    Scrollbars = DHTMLSafe1.Scrollbars
End Property

Public Property Let Scrollbars(ByVal New_Scrollbars As Boolean)
    DHTMLSafe1.Scrollbars() = New_Scrollbars
    PropertyChanged "Scrollbars"
End Property

Private Sub DHTMLSafe1_ShowContextMenu(ByVal xPos As Long, ByVal yPos As Long)
    RaiseEvent ShowContextMenu(xPos, yPos)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,ShowDetails
Public Property Get ShowDetails() As Boolean
Attribute ShowDetails.VB_Description = "property ShowDetails"
    ShowDetails = DHTMLSafe1.ShowDetails
End Property

Public Property Let ShowDetails(ByVal New_ShowDetails As Boolean)
    DHTMLSafe1.ShowDetails() = New_ShowDetails
    PropertyChanged "ShowDetails"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,ShowBorders
Public Property Get ShowBorders() As Boolean
Attribute ShowBorders.VB_Description = "property ShowBorders"
    ShowBorders = DHTMLSafe1.ShowBorders
End Property

Public Property Let ShowBorders(ByVal New_ShowBorders As Boolean)
    DHTMLSafe1.ShowBorders() = New_ShowBorders
    PropertyChanged "ShowBorders"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,SourceCodePreservation
Public Property Get SourceCodePreservation() As Boolean
Attribute SourceCodePreservation.VB_Description = "property SourceCodePreservation"
    SourceCodePreservation = DHTMLSafe1.SourceCodePreservation
End Property

Public Property Let SourceCodePreservation(ByVal New_SourceCodePreservation As Boolean)
    DHTMLSafe1.SourceCodePreservation() = New_SourceCodePreservation
    PropertyChanged "SourceCodePreservation"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = DHTMLSafe1.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    DHTMLSafe1.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,UseDivOnCarriageReturn
Public Property Get UseDivOnCarriageReturn() As Boolean
Attribute UseDivOnCarriageReturn.VB_Description = "property UseDivOnCarriageReturn"
    UseDivOnCarriageReturn = DHTMLSafe1.UseDivOnCarriageReturn
End Property

Public Property Let UseDivOnCarriageReturn(ByVal New_UseDivOnCarriageReturn As Boolean)
    DHTMLSafe1.UseDivOnCarriageReturn() = New_UseDivOnCarriageReturn
    PropertyChanged "UseDivOnCarriageReturn"
End Property

Private Sub DHTMLSafe1_Validate(Cancel As Boolean)
    RaiseEvent Validate(Cancel)
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    DHTMLSafe1.AbsoluteDropMode = PropBag.ReadProperty("AbsoluteDropMode", False)
    DHTMLSafe1.BaseURL = PropBag.ReadProperty("BaseURL", "")
    DHTMLSafe1.DocumentHTML = PropBag.ReadProperty("DocumentHTML", "")
    Set Palette = PropBag.ReadProperty("Palette", Nothing)
    UserControl.ScaleLeft = PropBag.ReadProperty("ScaleLeft", 0)
    UserControl.ScaleMode = PropBag.ReadProperty("ScaleMode", 1)
    UserControl.ScaleTop = PropBag.ReadProperty("ScaleTop", 0)
    UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 4800)
    DHTMLSafe1.ScrollbarAppearance = PropBag.ReadProperty("ScrollbarAppearance", 1)
    DHTMLSafe1.Scrollbars = PropBag.ReadProperty("Scrollbars", True)
    DHTMLSafe1.ShowDetails = PropBag.ReadProperty("ShowDetails", False)
    DHTMLSafe1.ShowBorders = PropBag.ReadProperty("ShowBorders", False)
    DHTMLSafe1.SourceCodePreservation = PropBag.ReadProperty("SourceCodePreservation", True)
    DHTMLSafe1.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    DHTMLSafe1.UseDivOnCarriageReturn = PropBag.ReadProperty("UseDivOnCarriageReturn", False)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    
    DHTMLSafe1.SourceCodePreservation = PropBag.ReadProperty("Locked", True)
    UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 3045)
    m_BrowseMode = PropBag.ReadProperty("BrowseMode", False)
    m_ShowToolbar = PropBag.ReadProperty("ShowToolbar", m_def_ShowToolbar)
    
End Sub

Private Sub UserControl_Show()
    showPrintValue "UserControl_Show()"
    If Ambient.UserMode = True Then
        CoolBar1.Visible = m_ShowToolbar
'        FlatCombo cmbFont
'        FlatCombo cmbFontSize
'        FlatCombo cmbParagraph
    End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    Call PropBag.WriteProperty("AbsoluteDropMode", DHTMLSafe1.AbsoluteDropMode, False)
    Call PropBag.WriteProperty("BaseURL", DHTMLSafe1.BaseURL, "")
    Call PropBag.WriteProperty("DocumentHTML", DHTMLSafe1.DocumentHTML, "")
    Call PropBag.WriteProperty("Palette", Palette, Nothing)
'    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 3600)
    Call PropBag.WriteProperty("ScaleLeft", UserControl.ScaleLeft, 0)
    Call PropBag.WriteProperty("ScaleMode", UserControl.ScaleMode, 1)
    Call PropBag.WriteProperty("ScaleTop", UserControl.ScaleTop, 0)
    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 4800)
    Call PropBag.WriteProperty("ScrollbarAppearance", DHTMLSafe1.ScrollbarAppearance, 1)
    Call PropBag.WriteProperty("Scrollbars", DHTMLSafe1.Scrollbars, True)
    Call PropBag.WriteProperty("ShowDetails", DHTMLSafe1.ShowDetails, False)
    Call PropBag.WriteProperty("ShowBorders", DHTMLSafe1.ShowBorders, False)
    Call PropBag.WriteProperty("SourceCodePreservation", DHTMLSafe1.SourceCodePreservation, True)
    Call PropBag.WriteProperty("ToolTipText", DHTMLSafe1.ToolTipText, "")
    Call PropBag.WriteProperty("UseDivOnCarriageReturn", DHTMLSafe1.UseDivOnCarriageReturn, False)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    
    Call PropBag.WriteProperty("Locked", DHTMLSafe1.SourceCodePreservation, True)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 3045)
    Call PropBag.WriteProperty("BrowseMode", m_BrowseMode, False)
    Call PropBag.WriteProperty("ShowToolbar", m_ShowToolbar, m_def_ShowToolbar)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,SourceCodePreservation
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "property SourceCodePreservation"
    Locked = DHTMLSafe1.SourceCodePreservation
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    DHTMLSafe1.SourceCodePreservation() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,DOM
'Public Property Get DOM() As IHTMLDocument
'    Set DOM = DHTMLSafe1.DOM
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    UserControl.ScaleHeight() = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,BrowseMode
Public Property Get BrowseMode() As Boolean
Attribute BrowseMode.VB_Description = "property BrowseMode"
    BrowseMode = m_BrowseMode
    'DHTMLSafe1.BrowseMode
End Property

Public Property Let BrowseMode(ByVal New_BrowseMode As Boolean)
    m_BrowseMode = New_BrowseMode
    DHTMLSafe1.BrowseMode = m_BrowseMode
    If Ambient.UserMode = True Then
        DisbaleToolbar m_BrowseMode
    End If
    PropertyChanged "BrowseMode"
End Property
Private Sub DisbaleToolbar(ByVal what As Boolean)
    what = Not what
    Dim i As Integer
    For i = 1 To Toolbar2.Buttons.Count
        If Toolbar2.Buttons(i).Style <> tbrSeparator Then
           Toolbar2.Buttons(i).Enabled = what
        End If
    Next
    cmbFont.Enabled = what
    cmbFontSize.Enabled = what
    cmbParagraph.Enabled = what
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,SaveDocument
Public Sub SaveDocument(pathIn As Variant, Optional promptUser As Variant)
Attribute SaveDocument.VB_Description = "method SaveDocument"
    DHTMLSafe1.SaveDocument pathIn, promptUser
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DHTMLSafe1,DHTMLSafe1,-1,PrintDocument
Public Sub PrintDocument(Optional withUI As Variant)
Attribute PrintDocument.VB_Description = "method PrintDocument"
    DHTMLSafe1.PrintDocument withUI
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowToolbar() As Boolean
Attribute ShowToolbar.VB_Description = "Show the Toolbar with the HTML Editor"
    ShowToolbar = m_ShowToolbar
End Property

Public Property Let ShowToolbar(ByVal New_ShowToolbar As Boolean)
    m_ShowToolbar = New_ShowToolbar
    CoolBar1.Visible = m_ShowToolbar
    PropertyChanged "ShowToolbar"
    UserControl_Resize
End Property

Private Sub showPrintValue(ByVal bb As String)
    Debug.Print bb & "User Mode : " & Ambient.UserMode
End Sub


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.UserControl AutoEntry 
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7920
   ScaleHeight     =   5115
   ScaleWidth      =   7920
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "UserControl1.ctx":0000
      Height          =   2475
      Left            =   180
      OleObjectBlob   =   "UserControl1.ctx":0014
      TabIndex        =   10
      Top             =   2220
      Width           =   6555
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6480
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":09E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":0AF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":0C53
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":0DAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":0F07
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":1061
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":11BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":1315
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":18AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":1E49
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":1FA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":20FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":2697
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":27F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":2D8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":2E9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":3437
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":39D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":3F6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":4505
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":4A9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":5039
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   953
      ButtonWidth     =   1349
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Add"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Delete"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancel"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Update"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Preview"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Exit"
            ImageIndex      =   19
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   3315
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   6795
      Begin VB.PictureBox WorkArea 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3360
         Left            =   60
         ScaleHeight     =   3360
         ScaleWidth      =   5070
         TabIndex        =   2
         Top             =   120
         Width           =   5070
         Begin VB.VScrollBar VScroll1 
            Height          =   2565
            Left            =   4635
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   15
            Width           =   240
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   240
            Left            =   0
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   2925
            Width           =   4740
         End
         Begin VB.PictureBox PicScroll 
            BorderStyle     =   0  'None
            Height          =   2670
            Left            =   60
            ScaleHeight     =   2670
            ScaleWidth      =   4485
            TabIndex        =   3
            Top             =   -15
            Width           =   4485
            Begin VB.PictureBox PicArea 
               BorderStyle     =   0  'None
               Height          =   2445
               Left            =   -60
               ScaleHeight     =   2445
               ScaleWidth      =   3240
               TabIndex        =   4
               Top             =   60
               Width           =   3240
               Begin VB.CommandButton cmdDate 
                  Caption         =   ".."
                  Height          =   360
                  Index           =   0
                  Left            =   2550
                  TabIndex        =   6
                  Top             =   285
                  Visible         =   0   'False
                  Width           =   330
               End
               Begin VB.TextBox Text1 
                  DataSource      =   "Data1"
                  Height          =   315
                  Index           =   0
                  Left            =   390
                  MultiLine       =   -1  'True
                  TabIndex        =   5
                  Top             =   300
                  Width           =   2115
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Label1"
                  Height          =   195
                  Index           =   0
                  Left            =   60
                  TabIndex        =   7
                  Top             =   0
                  Width           =   1905
               End
            End
         End
         Begin VB.Data Data1 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   180
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   3480
            Visible         =   0   'False
            Width           =   2415
         End
      End
   End
End
Attribute VB_Name = "AutoEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim IdxFocus As Integer
Private Declare Function ShowScrollBar Lib "User32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
'Default Property Values:
Const m_def_DatabaseFileName = ""
Const m_def_SourceTable = ""
'Property Variables:
Dim m_DatabaseFileName As String
Dim m_SourceTable As String
'Event Declarations:
Event Change() 'MappingInfo=Text1(0),Text1,0,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event AddClick()
Event EditClick()
Event CancelClick()
Event UpdateClick()
Event ExitClick()
Event PreviewClick()
Event DeleteClick()



Private Function GetFieldWidth(rnType As Integer)
  'determines the form control width
  'based on the field type
  Select Case rnType
    Case dbBoolean
      GetFieldWidth = 850
    Case dbByte
      GetFieldWidth = 650
    Case dbInteger
      GetFieldWidth = 900
    Case dbLong
      GetFieldWidth = 1100
    Case dbCurrency
      GetFieldWidth = 1800
    Case dbSingle
      GetFieldWidth = 1800
    Case dbDouble
      GetFieldWidth = 2200
    Case dbDate
      GetFieldWidth = 1000
    Case dbText
      GetFieldWidth = 3250
    Case dbMemo
      GetFieldWidth = 3250
    Case Else
      GetFieldWidth = 3250
  End Select

End Function



Private Sub HScroll1_Change()
    PicArea.Left = -HScroll1.Value
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Dim Nilai As Long
With Text1(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
    If .Top + .Height >= PicScroll.Height And IdxFocus < Index Then
       Nilai = VScroll1.Value + VScroll1.SmallChange
       If Nilai > VScroll1.Max Then Nilai = VScroll1.Max
       VScroll1.Value = Nilai
    ElseIf PicArea.Top + .Top < 0 And IdxFocus > Index Then
       Nilai = VScroll1.Value - VScroll1.SmallChange
       If Nilai < VScroll1.Min Then Nilai = VScroll1.Min
       If Nilai > VScroll1.Min Then Nilai = VScroll1.Min
       VScroll1.Value = Nilai
    End If
End With
IdxFocus = Index
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Index = Index + 1
        If Index > Text1.Count - 1 Then Index = 0
        Text1(Index).SetFocus
     End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            RaiseEvent AddClick
        Case 2
            RaiseEvent EditClick
        Case 3
            RaiseEvent DeleteClick
        Case 4
            RaiseEvent CancelClick
        Case 5
            RaiseEvent UpdateClick
        Case 6
            RaiseEvent PreviewClick
        Case 7
            RaiseEvent ExitClick
    End Select
End Sub

Private Sub UserControl_Resize()
On Local Error Resume Next
Data1.DatabaseName = DatabaseFileName
Data1.RecordSource = SourceTable
Data1.Refresh
    
    Dim Jenis As Integer, Ukuran As Integer, JumlahField As Integer, WidthText  As Long, SaveWidthText As Long
    JumlahField = Data1.Recordset.Fields.Count
    For k = 0 To JumlahField - 1
        If k > 0 Then
           Load Text1(k)
           Load Label1(k)
        End If
        Text1(k).DataField = Data1.Recordset.Fields(k).Name
        Label1(k).Caption = Text1(k).DataField & " :"
        Jenis = Data1.Recordset.Fields(k).Type
        Ukuran = Data1.Recordset.Fields(k).Size
        Text1(k).Tag = Jenis
        Select Case Jenis
            Case dbDate
                Ukuran = Ukuran + 2
                If k > 0 Then Load cmdDate(k)
            Case dbInteger, dbLong
                Text1(k).Alignment = 1
        Case dbMemo
           Text1(k).Height = 1000
           ShowScrollBar Text1(k).hWnd, 1, True
        End Select
        
        WidthText = GetFieldWidth(Jenis)
        Text1(k).Width = WidthText
        Text1(k).MaxLength = Ukuran
        Text1(k).Left = Label1(k).Left + Label1(k).Width + 80
        If k > 0 Then
            Text1(k).Top = Text1(k - 1).Top + Text1(k - 1).Height + 80
        End If
        If SaveWidthText < WidthText + Text1(k).Left + 100 Then SaveWidthText = WidthText + Text1(k).Left + 100
        Label1(k).Top = Text1(k).Top
        
        If Jenis = dbDate Then
            cmdDate(k).Left = Text1(k).Left + Text1(k).Width + 40
            cmdDate(k).Top = Text1(k).Top
            cmdDate(k).Height = Text1(k).Height
            cmdDate(k).Visible = True
        End If
        
        Text1(k).Visible = True
        Label1(k).Visible = True
    Next
    
    Frame1.Width = Width - Frame1.Left
    DBGrid1.Top = Height - DBGrid1.Height + 60
    DBGrid1.Height = Height - DBGrid1.Top
    DBGrid1.Width = Frame1.Width
    DBGrid1.Left = Frame1.Left
    Frame1.Height = DBGrid1.Top - Frame1.Top - 60
    WorkArea.Left = 30
    WorkArea.Top = 120
    WorkArea.Width = Frame1.Width - 60
    WorkArea.Height = Frame1.Height - 150
    PicScroll.Left = 0
    PicScroll.Top = 0
    PicScroll.Width = WorkArea.Width - VScroll1.Width - 60
    PicScroll.Height = WorkArea.Height - HScroll1.Height - 60
    PicArea.Top = 0
    PicArea.Left = 0
    PicArea.Height = Text1(0).Top + (Text1(JumlahField - 1).Top + Text1(JumlahField - 1).Height)
    PicArea.Width = SaveWidthText
    VScroll1.Left = PicScroll.Left + PicScroll.Width + 60
    VScroll1.Top = PicScroll.Top
    VScroll1.Height = PicScroll.Height + 60
    
    HScroll1.Left = PicScroll.Left
    HScroll1.Top = PicScroll.Top + PicScroll.Height + 60
    HScroll1.Width = PicScroll.Width + 60
    
    If PicArea.Height > PicScroll.Height Then
       VScroll1.Max = PicArea.Height - PicScroll.Height
       VScroll1.LargeChange = PicArea.Height / 2
       VScroll1.SmallChange = PicArea.Height / JumlahField - 1
       VScroll1.Visible = True
       VScroll1.Enabled = True
    Else
       'VScroll1.Visible = False
       VScroll1.Enabled = False
    End If
    
    If PicArea.Width > PicScroll.Width Then
       HScroll1.Max = PicArea.Width - PicScroll.Width
       HScroll1.LargeChange = PicArea.Width / 2
       HScroll1.SmallChange = PicArea.Width / 10
       HScroll1.Visible = True
       HScroll1.Enabled = True
    Else
       HScroll1.Enabled = False
    End If
    
    If HScroll1.Enabled = False Then
       VScroll1.Height = VScroll1.Height + HScroll1.Height
       PicScroll.Height = PicScroll.Height + HScroll1.Height
       HScroll1.Visible = False
    End If
   
    If VScroll1.Enabled = False Then
       HScroll1.Width = HScroll1.Width + VScroll1.Width
       PicScroll.Width = PicScroll.Width + VScroll1.Width
       VScroll1.Visible = False
    End If
    
    
End Sub

Private Sub UserControl_Show()
    UserControl_Resize
End Sub

Private Sub VScroll1_Change()
    PicArea.Top = -VScroll1.Value
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get DatabaseFileName() As String
    DatabaseFileName = m_DatabaseFileName
End Property

Public Property Let DatabaseFileName(ByVal New_DatabaseFileName As String)
    m_DatabaseFileName = New_DatabaseFileName
    PropertyChanged "DatabaseFileName"
    Data1.DatabaseName = DatabaseFileName
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get SourceTable() As String
    SourceTable = m_SourceTable
End Property

Public Property Let SourceTable(ByVal New_SourceTable As String)
    m_SourceTable = New_SourceTable
    PropertyChanged "SourceTable"
    Data1.RecordSource = SourceTable
    Data1.Refresh
    UserControl_Resize
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_DatabaseFileName = m_def_DatabaseFileName
    m_SourceTable = m_def_SourceTable
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_DatabaseFileName = PropBag.ReadProperty("DatabaseFileName", m_def_DatabaseFileName)
    m_SourceTable = PropBag.ReadProperty("SourceTable", m_def_SourceTable)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("DatabaseFileName", m_DatabaseFileName, m_def_DatabaseFileName)
    Call PropBag.WriteProperty("SourceTable", m_SourceTable, m_def_SourceTable)
End Sub

Private Sub Text1_Change(Index As Integer)
    RaiseEvent Change
End Sub


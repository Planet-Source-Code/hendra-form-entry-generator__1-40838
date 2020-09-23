VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto Entry"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin Project1.AutoEntry AutoEntry1 
      Height          =   5235
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   9234
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    AutoEntry1.DatabaseFileName = App.Path & "\Nwind.mdb"
    AutoEntry1.SourceTable = "Customers"
End Sub

Private Sub AutoEntry1_ExitClick()
    Unload Me
End Sub

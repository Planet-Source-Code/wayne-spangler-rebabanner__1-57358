VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RebaBanner Help"
   ClientHeight    =   11970
   ClientLeft      =   1350
   ClientTop       =   1800
   ClientWidth     =   13320
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11970
   ScaleWidth      =   13320
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   4290
      Picture         =   "frmHelp.frx":0CCA
      ScaleHeight     =   600
      ScaleWidth      =   4680
      TabIndex        =   1
      Top             =   11265
      Width           =   4740
   End
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   2250
      Left            =   165
      TabIndex        =   0
      Top             =   165
      Width           =   4500
      ExtentX         =   7937
      ExtentY         =   3969
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_HelpFile As String

Private Sub Form_Load()
    WB.Navigate m_HelpFile
End Sub

Private Sub Form_Resize()
    WB.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200 - Picture1.Height
End Sub

Public Property Let HelpFile(ByVal vNewValue As String)
    m_HelpFile = vNewValue
End Property

Private Sub Picture1_Click()
    Unload Me
End Sub

VERSION 5.00
Begin VB.Form frmCustomSize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Size"
   ClientHeight    =   2055
   ClientLeft      =   6735
   ClientTop       =   3735
   ClientWidth     =   3780
   Icon            =   "frmCustomSize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   435
      TabIndex        =   5
      Top             =   1335
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2220
      TabIndex        =   4
      Top             =   1350
      Width           =   1215
   End
   Begin VB.TextBox txtCustomX 
      Height          =   300
      Left            =   2160
      TabIndex        =   3
      Top             =   825
      Width           =   1260
   End
   Begin VB.TextBox txtCustomY 
      Height          =   300
      Left            =   2160
      TabIndex        =   1
      Top             =   390
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Width of Banner"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   300
      TabIndex        =   2
      Top             =   825
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Height of Banner"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   300
      TabIndex        =   0
      Top             =   390
      Width           =   1605
   End
End
Attribute VB_Name = "frmCustomSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'************************************************************************
'************************************************************************
Private Sub Command1_Click()
    CustomX = txtCustomX
    CustomY = txtCustomY
    Unload Me
End Sub

'************************************************************************
'************************************************************************
Private Sub Command2_Click()
    CustomX = 0: CustomY = 0
    Unload Me
End Sub

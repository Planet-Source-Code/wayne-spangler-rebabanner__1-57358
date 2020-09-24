VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3165
   ClientLeft      =   3195
   ClientTop       =   3150
   ClientWidth     =   7560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   211
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   504
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   960
      Left            =   240
      Picture         =   "frmSplash.frx":0CCA
      ScaleHeight     =   900
      ScaleWidth      =   7020
      TabIndex        =   0
      Top             =   570
      Width           =   7080
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      X1              =   496
      X2              =   496
      Y1              =   198
      Y2              =   15
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      X1              =   8
      X2              =   496
      Y1              =   200
      Y2              =   199
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      X1              =   9
      X2              =   496
      Y1              =   14
      Y2              =   14
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      X1              =   9
      X2              =   8
      Y1              =   13
      Y2              =   199
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Version 1.0.0"
      Height          =   195
      Left            =   3315
      TabIndex        =   3
      Top             =   1770
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "By"
      Height          =   195
      Left            =   3690
      TabIndex        =   2
      Top             =   2025
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Rebaware"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3165
      TabIndex        =   1
      Top             =   2295
      Width           =   1230
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


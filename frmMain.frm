VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   Caption         =   "Banner"
   ClientHeight    =   5940
   ClientLeft      =   1305
   ClientTop       =   2175
   ClientWidth     =   8400
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   396
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   560
   Begin TabDlg.SSTab SSTab1 
      Height          =   4260
      Left            =   1485
      TabIndex        =   3
      Top             =   1200
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   7514
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   529
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmMain.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "BannerSizes"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "Label6"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Text"
      TabPicture(1)   =   "frmMain.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "OutlineColor"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "ChangeFontSize"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txt1FontSize"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txt1Italic"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Text1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame4"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Shadow"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtShadow"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "OutlineText"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Command4"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Frame3"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmdFontColor"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cboFontName"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txt1Bold"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txt1Underline"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "Picture"
      TabPicture(2)   =   "frmMain.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DisplayPic"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdBrowsePic"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame7"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "MovePicture"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Border"
      TabPicture(3)   =   "frmMain.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).Control(1)=   "ShowBorder"
      Tab(3).Control(2)=   "LineBorder"
      Tab(3).Control(3)=   "ShadeBorder"
      Tab(3).Control(4)=   "Frame5"
      Tab(3).Control(5)=   "BorderWidth"
      Tab(3).Control(6)=   "Label10"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "Buttons"
      TabPicture(4)   =   "frmMain.frx":0D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command3"
      Tab(4).Control(1)=   "Command2"
      Tab(4).Control(2)=   "Command1"
      Tab(4).Control(3)=   "Label21"
      Tab(4).Control(4)=   "Label20"
      Tab(4).Control(5)=   "Image2"
      Tab(4).Control(6)=   "Image1"
      Tab(4).Control(7)=   "Label19"
      Tab(4).Control(8)=   "Label18"
      Tab(4).Control(9)=   "Label17"
      Tab(4).Control(10)=   "Label9"
      Tab(4).ControlCount=   11
      Begin VB.CheckBox txt1Underline 
         Caption         =   "Underline"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2025
         TabIndex        =   72
         Top             =   2310
         Width           =   1170
      End
      Begin VB.CheckBox txt1Bold 
         Caption         =   "Bold"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   71
         Top             =   2295
         Width           =   735
      End
      Begin VB.ComboBox cboFontName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmMain.frx":0D56
         Left            =   240
         List            =   "frmMain.frx":0D58
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   1395
         Width           =   1815
      End
      Begin VB.CommandButton cmdFontColor 
         Caption         =   "&Font Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   69
         Top             =   1815
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Caption         =   "Move Text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   240
         TabIndex        =   62
         Top             =   3000
         Width           =   2535
         Begin VB.TextBox txtVert 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   405
            TabIndex        =   64
            Text            =   "8"
            Top             =   495
            Width           =   375
         End
         Begin VB.TextBox txtHorz 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1365
            TabIndex        =   63
            Text            =   "8"
            Top             =   495
            Width           =   495
         End
         Begin MSComCtl2.UpDown VertMove 
            Height          =   405
            Left            =   765
            TabIndex        =   65
            Top             =   495
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   714
            _Version        =   393216
            Value           =   25
            BuddyControl    =   "txtVert"
            BuddyDispid     =   196614
            OrigLeft        =   600
            OrigTop         =   600
            OrigRight       =   1215
            OrigBottom      =   915
            Increment       =   2
            Max             =   500
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown HorzMove 
            Height          =   285
            Left            =   1845
            TabIndex        =   66
            Top             =   495
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtHorz"
            BuddyDispid     =   196615
            OrigLeft        =   1920
            OrigTop         =   600
            OrigRight       =   2520
            OrigBottom      =   915
            Increment       =   5
            Max             =   50
            Orientation     =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vertical"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   405
            TabIndex        =   68
            Top             =   255
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Horzontal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1245
            TabIndex        =   67
            Top             =   255
            Width           =   675
         End
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00000000&
         Caption         =   "Outline Color"
         Height          =   285
         Left            =   4905
         TabIndex        =   60
         Top             =   1770
         Width           =   1275
      End
      Begin VB.CheckBox OutlineText 
         Caption         =   "Outline Text"
         Height          =   345
         Left            =   3615
         TabIndex        =   59
         Top             =   1710
         Width           =   1200
      End
      Begin VB.CommandButton MovePicture 
         Caption         =   "Move Picture to Top"
         Height          =   375
         Left            =   -73080
         TabIndex        =   58
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Height          =   585
         Left            =   -70650
         Picture         =   "frmMain.frx":0D5A
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1845
         Width           =   1440
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command1"
         Height          =   495
         Left            =   -72780
         Picture         =   "frmMain.frx":2D94
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1830
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   -74580
         TabIndex        =   49
         Top             =   1815
         Width           =   1575
      End
      Begin VB.ComboBox BannerSizes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmMain.frx":3D0D
         Left            =   -72600
         List            =   "frmMain.frx":3D0F
         Style           =   2  'Dropdown List
         TabIndex        =   47
         ToolTipText     =   "Banner Sizes"
         Top             =   840
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         Height          =   3015
         Left            =   -70980
         TabIndex        =   40
         Top             =   600
         Width           =   2055
         Begin VB.Image LBorderLighter 
            Height          =   375
            Left            =   1440
            Picture         =   "frmMain.frx":3D11
            Stretch         =   -1  'True
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Light Border"
            Height          =   240
            Left            =   120
            TabIndex        =   46
            Top             =   720
            Width           =   1080
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Dark Border"
            Height          =   240
            Left            =   120
            TabIndex        =   45
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Image LBorderDarker 
            Height          =   375
            Left            =   1440
            Picture         =   "frmMain.frx":4153
            Stretch         =   -1  'True
            Top             =   840
            Width           =   375
         End
         Begin VB.Image DBorderLighter 
            Height          =   375
            Left            =   1440
            Picture         =   "frmMain.frx":4595
            Stretch         =   -1  'True
            Top             =   1920
            Width           =   375
         End
         Begin VB.Image DBorderDarker 
            Height          =   375
            Left            =   1440
            Picture         =   "frmMain.frx":49D7
            Stretch         =   -1  'True
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label Label13 
            Caption         =   "Darker"
            Height          =   255
            Left            =   1320
            TabIndex        =   44
            Top             =   2640
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "Lighter"
            Height          =   255
            Left            =   1320
            TabIndex        =   43
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "Darker"
            Height          =   255
            Left            =   1320
            TabIndex        =   42
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "Lighter"
            Height          =   255
            Left            =   1320
            TabIndex        =   41
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.CheckBox ShowBorder 
         Caption         =   "Show"
         Height          =   255
         Left            =   -73200
         TabIndex        =   37
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton LineBorder 
         Caption         =   "Line"
         Height          =   255
         Left            =   -74280
         TabIndex        =   36
         Top             =   2040
         Width           =   735
      End
      Begin VB.OptionButton ShadeBorder 
         Caption         =   "Shade"
         Height          =   255
         Left            =   -74280
         TabIndex        =   35
         Top             =   2400
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.Frame Frame5 
         Height          =   1215
         Left            =   -73320
         TabIndex        =   34
         Top             =   1680
         Visible         =   0   'False
         Width           =   2175
         Begin VB.CommandButton cmdBorderColor 
            Caption         =   "Border Color"
            Height          =   375
            Left            =   360
            TabIndex        =   39
            Top             =   465
            Width           =   1455
         End
      End
      Begin VB.TextBox BorderWidth 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   -73200
         TabIndex        =   33
         Text            =   "1"
         Top             =   1080
         Width           =   495
      End
      Begin VB.Frame Frame7 
         Caption         =   "Move"
         Height          =   1455
         Left            =   -71760
         TabIndex        =   30
         Top             =   1440
         Width           =   2415
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   1560
            TabIndex        =   31
            Text            =   "5"
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label8 
            Caption         =   "Move Distance"
            Height          =   375
            Left            =   1560
            TabIndex        =   32
            Top             =   360
            Width           =   735
         End
         Begin VB.Image MoveRight 
            Height          =   480
            Left            =   840
            Picture         =   "frmMain.frx":4E19
            Top             =   600
            Width           =   480
         End
         Begin VB.Image MoveLeft 
            Height          =   480
            Left            =   120
            Picture         =   "frmMain.frx":525B
            Top             =   600
            Width           =   480
         End
         Begin VB.Image MoveUP 
            Height          =   480
            Left            =   480
            Picture         =   "frmMain.frx":569D
            Top             =   360
            Width           =   480
         End
         Begin VB.Image MoveDown 
            Height          =   480
            Left            =   480
            Picture         =   "frmMain.frx":5ADF
            Top             =   840
            Width           =   480
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Resize"
         Height          =   1455
         Left            =   -74640
         TabIndex        =   25
         Top             =   1440
         Width           =   2415
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   1680
            TabIndex        =   26
            Text            =   "3"
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "Resize Amount"
            Height          =   495
            Left            =   1560
            TabIndex        =   29
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Smaller"
            Height          =   240
            Left            =   720
            TabIndex        =   28
            Top             =   960
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Larger"
            Height          =   240
            Left            =   720
            TabIndex        =   27
            Top             =   480
            Width           =   585
         End
         Begin VB.Image ResizeDown 
            Height          =   480
            Left            =   120
            Picture         =   "frmMain.frx":5F21
            Top             =   840
            Width           =   480
         End
         Begin VB.Image ResizeUp 
            Height          =   480
            Left            =   120
            Picture         =   "frmMain.frx":6363
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.CommandButton cmdBrowsePic 
         Caption         =   "Load Picture"
         Height          =   375
         Left            =   -72480
         TabIndex        =   24
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox DisplayPic 
         Caption         =   "Display Picture"
         Enabled         =   0   'False
         Height          =   240
         Left            =   -72600
         TabIndex        =   23
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtShadow 
         Height          =   360
         Left            =   4905
         TabIndex        =   20
         Text            =   "1"
         Top             =   1110
         Width           =   375
      End
      Begin VB.CheckBox Shadow 
         Caption         =   "Shadow"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3645
         TabIndex        =   19
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Caption         =   "Alignment"
         Height          =   1695
         Left            =   3510
         TabIndex        =   12
         Top             =   2325
         Width           =   2895
         Begin VB.CommandButton cmdTop 
            Caption         =   "&Top"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdMid 
            Caption         =   "&Mid"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton cmdBottom 
            Caption         =   "&Bottom"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton cmdLeft 
            Caption         =   "&Left"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   15
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdCenter 
            Caption         =   "&Center"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   14
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton cmdRight 
            Caption         =   "&Right"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   13
            Top             =   1080
            Width           =   1095
         End
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   360
         Width           =   6555
      End
      Begin VB.CheckBox txt1Italic 
         Caption         =   "Italic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1215
         TabIndex        =   10
         Top             =   2295
         Width           =   825
      End
      Begin VB.TextBox txt1FontSize 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2325
         TabIndex        =   9
         Text            =   "18"
         Top             =   1425
         Width           =   375
      End
      Begin VB.Frame Frame1 
         Caption         =   "Background"
         Height          =   1455
         Left            =   -73560
         TabIndex        =   4
         Top             =   1680
         Width           =   3240
         Begin VB.OptionButton optColor 
            Caption         =   "Color"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   480
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optGraphic 
            Caption         =   "Graphic"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdSelectColor 
            Caption         =   "Select Color"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1710
            TabIndex        =   6
            Top             =   345
            Width           =   1335
         End
         Begin VB.CommandButton cmdBrowseBackground 
            Caption         =   "Browse"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1710
            TabIndex        =   5
            Top             =   945
            Width           =   1335
         End
      End
      Begin MSComCtl2.UpDown ChangeFontSize 
         Height          =   405
         Left            =   2685
         TabIndex        =   21
         Top             =   1425
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   714
         _Version        =   393216
         Value           =   18
         BuddyControl    =   "txt1FontSize"
         BuddyDispid     =   196668
         OrigLeft        =   656
         OrigTop         =   248
         OrigRight       =   672
         OrigBottom      =   273
         Increment       =   2
         Max             =   60
         Min             =   8
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Font"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   73
         Top             =   1035
         Width           =   435
      End
      Begin VB.Label OutlineColor 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   6330
         TabIndex        =   61
         Top             =   1740
         Width           =   315
      End
      Begin VB.Label Label21 
         Caption         =   "Downside: Image control can not get the focus so you can't tab to it."
         Height          =   825
         Left            =   -70305
         TabIndex        =   57
         Top             =   2880
         Width           =   1350
      End
      Begin VB.Label Label20 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Command button or Image Control - Your choice. Click on button to see it work."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   540
         Left            =   -74670
         TabIndex        =   56
         Top             =   495
         Width           =   3780
      End
      Begin VB.Image Image2 
         Height          =   705
         Left            =   -72645
         Picture         =   "frmMain.frx":67A5
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   2160
      End
      Begin VB.Image Image1 
         Height          =   510
         Left            =   -74595
         Picture         =   "frmMain.frx":EC53
         Stretch         =   -1  'True
         Top             =   3090
         Width           =   1530
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Graphical Button - Image control - 2 buttons, Same pictures, Different sizes."
         Height          =   420
         Left            =   -74625
         TabIndex        =   55
         Top             =   2625
         Width           =   3540
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Windows Button with Pictured Background. As you can see the button needs to be resized so it can all be seen."
         Height          =   1155
         Left            =   -70665
         TabIndex        =   54
         Top             =   495
         Width           =   1620
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Windows Button with Textured Background"
         Height          =   420
         Left            =   -72795
         TabIndex        =   52
         Top             =   1215
         Width           =   1620
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Normal Windows Button"
         Height          =   360
         Left            =   -74520
         TabIndex        =   50
         Top             =   1215
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Banner Sizes"
         Height          =   195
         Left            =   -73680
         TabIndex        =   48
         Top             =   960
         Width           =   945
      End
      Begin VB.Label Label10 
         Caption         =   "Width"
         Height          =   255
         Left            =   -72600
         TabIndex        =   38
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Font Size"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2265
         TabIndex        =   22
         Top             =   1080
         Width           =   825
      End
   End
   Begin VB.PictureBox Display 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   120
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   87
      TabIndex        =   1
      ToolTipText     =   "The Banner Display"
      Top             =   120
      Width           =   1305
   End
   Begin VB.PictureBox PicBG 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3540
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4650
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox BanPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   75
      Left            =   2325
      Picture         =   "frmMain.frx":17101
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   210
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu hyp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu hyp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp1 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'************************************************************************
' When the bannersizes is click then a size has been selected. Call the
' SetSize routine and then redraw the banner.
'************************************************************************
Private Sub BannerSizes_Click()
    If BannerSizes.Text = "- - Custom - -" Then
        With frmCustomSize     'the custom size dialog
        'set the values to the current custom values
            .txtCustomX = CustomX
            .txtCustomY = CustomY
            ' get new values
            .Show 1
        End With
        'if either come back a 0 we can not use so exit doing nothing
        If CustomX = 0 Or CustomY = 0 Then Exit Sub
        'otherwise set the size according to the custom size
        SetSize Str(CustomX) & " x " & Str(CustomY)
    Else
        'set the banner size according to the banner size list
        SetSize BannerSizes.List(BannerSizes.ListIndex)
    End If
    Redraw
    Dirty = True
End Sub

'************************************************************************
' Sets the width of the border around the banner
'************************************************************************
Private Sub BorderWidth_Change()
    Dirty = True
    Redraw
End Sub

'************************************************************************
' Changes the color of the border if the border is set to line and width of
' the line is 1 or greater
' Has no effect when shade is selected
'************************************************************************
Private Sub cmdBorderColor_Click()
    'chooses the color for the border using the common dialog
    CommonDialog1.ShowColor
    'set the border color to that which was chosen
    BorderColor = CommonDialog1.color
    'and again redraw the banner
    Redraw
    Dirty = True
End Sub

'************************************************************************
' sets the text to the bottom of the picture box.
' if the border is displayed then it is moved to above the border
'************************************************************************
Private Sub cmdBottom_Click()
    txtVert = Display.ScaleHeight - Display.TextHeight(Text1)
    If ShowBorder = True Then
        txtVert = txtVert - Val(BorderWidth)
    End If
    Dirty = True
End Sub

'************************************************************************
' This button calls the common dialog box to select a graphic. It is then
' loaded into the BackgroundGraphic imagebox to be used if needed in the
' banner.
'************************************************************************
Private Sub cmdBrowseBackground_Click()
    With CommonDialog1
        'The filter selects just bmp,jpg and gif files
        .Filter = "bmp,jpg,gif,wmf|*.bmp;*.jpg;*.gif;*.wmf"
        .ShowOpen
        'When it returns you have to check that a valid file was returned.
        If .FileName <> "" Then
            'if there is a file name then load it into the background graphic
            PicBG.Picture = LoadPicture(.FileName)
            'and save the file name
            BackgroundPic = .FileName
        End If
    End With
    'and redraw
    Redraw
End Sub

'************************************************************************
' centers the text in the picture box - left to right
' if there is a picture then center to the left of picture
'************************************************************************
Private Sub cmdCenter_Click()
    Dim pWidth As Long
    pWidth = 0
    If DisplayPic.Value = 1 Then
        pWidth = PictureWidth
    End If
    If ShowBorder = 1 Then
        pWidth = pWidth '+ Val(BorderWidth) + Val(BorderWidth)
    End If
    'subtract the width of the text from the display width and divide by 2
    txtHorz = (Display.ScaleWidth - Display.TextWidth(Text1)) / 2 + pWidth / 2
    Dirty = True
End Sub

'************************************************************************
' Selects to color of the font
'************************************************************************
Private Sub cmdFontColor_Click()
    With CommonDialog1
        .ShowColor
        Display.ForeColor = .color
    End With
    Redraw
    Dirty = True
End Sub

'************************************************************************
' moves the text to the left side of the banner
' if the border is set then the text is moved to the right of the border
' if the picture is set the the text is moved to the right of the picture
'************************************************************************
Private Sub cmdLeft_Click()
    'Just set the horzontal position to 0 + the width of the border
    txtHorz = 0
    'if the border is set add the border width to the left position
    If ShowBorder = 1 Then
        txtHorz = Val(BorderWidth)
    End If
    'if the picture is set add the width and left of the picture to the left position
    If DisplayPic.Value = 1 Then
        txtHorz = txtHorz + PictureWidth
    End If
    Dirty = True
End Sub

'************************************************************************
' Centers the text in the banner - top to bottom
'************************************************************************
Private Sub cmdMid_Click()
    txtVert = (Display.ScaleHeight - Display.TextHeight("x")) / 2
    Dirty = True
End Sub

'************************************************************************
' Moves the text to the right side of banner
' if the border is set then subtract this from the position
'************************************************************************
Private Sub cmdRight_Click()
    'subtract the text width from the display width
    txtHorz = Display.ScaleWidth - Display.TextWidth(Text1)
    If ShowBorder = True Then
        txtHorz = txtVert - Val(BorderWidth)
    End If
    Dirty = True
End Sub

'************************************************************************
' This button calls the common dialog box to select a color. It then
' sets the display backcolor and the the banner is redrawn.
'************************************************************************
Private Sub cmdSelectColor_Click()
    With CommonDialog1
        .ShowColor
        Display.BackColor = CommonDialog1.color
    End With
    Redraw
    Dirty = True
End Sub

'************************************************************************
' moves the text to the top of the banner
' if the border is set then positioned below the border
'************************************************************************
Private Sub cmdTop_Click()
    txtVert = 0
    If ShowBorder = True Then
        txtVert = txtVert + Val(BorderWidth)
    End If
    Dirty = True
End Sub

Private Sub Command4_Click()
    With CommonDialog1
        .ShowColor
        OutlineColor.BackColor = .color
    End With
    Redraw
    Dirty = True
End Sub

'************************************************************************
' Makes the dark side of the border darker
'************************************************************************
Private Sub DBorderDarker_Click()
    DarkDiff = DarkDiff - 0.05
    Redraw
    Dirty = True
End Sub

'************************************************************************
' Makes the dark side of the border lighter
'************************************************************************
Private Sub DBorderLighter_Click()
    'Lighten the dark side
    DarkDiff = DarkDiff + 0.05
    Redraw
    Dirty = True
End Sub

'************************************************************************
' When the form is resized the sstab control is kept at the bottom right
'************************************************************************
Private Sub Form_Resize()
    SSTab1.Move Me.ScaleWidth - SSTab1.Width - 8, Me.ScaleHeight - SSTab1.Height - 8
End Sub

'************************************************************************
' Exits the program
' checks if current banner needs to be saved.
'************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    Dim ans As Integer
    If Dirty Then
        ans = MsgBox("Save File?", vbYesNo, "File has not been saved")
        If ans = vbYes Then mnuSave_Click
    End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image1.Picture = LoadPicture(App.Path & "\homedown.bmp")
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image1.Picture = LoadPicture(App.Path & "\home.bmp")
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image2.Picture = LoadPicture(App.Path & "\homedown.bmp")
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image2.Picture = LoadPicture(App.Path & "\home.bmp")
End Sub

'************************************************************************
' Make the light side of border darker
'************************************************************************
Private Sub LBorderDarker_Click()
    'Darkens the light border
    LiteDiff = LiteDiff - 0.05
    'and redraw
    Redraw
    Dirty = True
End Sub

'************************************************************************
' Make the light side of border lighter
'************************************************************************
Private Sub LBorderLighter_Click()
    'This makes the lighter sides lighter
    LiteDiff = LiteDiff + 0.05
    'as usual redraw the banner
    Redraw
    Dirty = True
End Sub

'************************************************************************
' at the start of the program we have to load the font's on the computer
' this may take some time if there are a lot of fonts, so to kill some time
' we show a splash screen. I have the time set for 3 seconds to allow time for
' the splash screen to be read. In most cases and time to load the fonts.
'************************************************************************
Private Sub Form_Load()
    Dim Prog As String
    frmSplash.Show
    'get all the fonts and put them in a combo box
     Dim i As Integer
     Dim t As Long
     t = Timer + 3
     With cboFontName
       For i = 0 To Screen.FontCount - 1
          .AddItem Screen.Fonts(i)
          DoEvents
       Next i
       ' Set ListIndex to 0.
       .ListIndex = 0
    End With
    'When the form loads we want to add the banner sizes to the
    'Bannersizes listbox.
    With BannerSizes
        .AddItem "- - Custom - -"
        .AddItem "468 x 60"
        .AddItem "392 x 72"
        .AddItem "312 x 40"
        .AddItem "234 x 60"
        .AddItem "200 x 26"
        .AddItem "150 x 75"
        .AddItem "150 x 150"
        .AddItem "125 x 125"
        .AddItem "120 x 90"
        .AddItem "120 x 60"
        .AddItem "100 x 30"
        .AddItem "88 x 31"
        'Set the index to point to the first banner
        .ListIndex = 1
        'Then call the SetSize routine to resize the banner
        SetSize .List(.ListIndex)
    End With
    'set the percent to increase the light color and dark color for the borders
    LiteDiff = 1.1
    DarkDiff = 0.9
    ' Check is a file is in the command section
    If Command$ <> "" Then
        Prog = Command$
        ' Make sure it is a banner file
        If InStr(Prog, ".ban") > 0 Then
            ' If it is then get rid of the ban portion
            FileName = Left$(Prog, InStr(Prog, ".ban"))
            ' Check to see if we have a quote at the begining of the name
            ' If so remove it or we get an error trying to load file
            If Left$(FileName, 1) = Chr$(34) Then
                FileName = Mid$(FileName, 2)
            End If
            ' Now open the file
            OpenFile
        End If
    End If
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    ' if loading this takes less time then we wanted then wait for time to elapse
   ' Do While t > Timer
   '     DoEvents
   ' Loop
    ' now we can get rid of the splash screen and set dirty to false
    Unload frmSplash
    Dirty = False
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

'************************************************************************
' exit was selected from the menu so unload the program
'************************************************************************
Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuHelp_Click()
    frmHelp.HelpFile = App.Path & "\RebaBanner.htm"
    frmHelp.Show
End Sub

'************************************************************************
' Start a new banner. Resets some values
'************************************************************************
Private Sub mnuNew_Click()
    Dim ans As Integer
    ' If changes made to current then ask if want to be saved
    If Dirty Then
        ans = MsgBox("Save File?", vbYesNo, "File has not been saved")
        If ans = vbYes Then mnuSave_Click
    End If
    ' reset some values
    BannerSizes.ListIndex = 1
    optColor.Value = True
    Text1 = ""
    DisplayPic.Value = 0
    ShowBorder = 0
    FileName = ""
    Dirty = False
End Sub

'************************************************************************
' Opens an existing file
'************************************************************************
Private Sub mnuOpen_Click()
    Dim ans As Integer
    ' check if any changes made to current banner
    If Dirty Then
        ans = MsgBox("Save File?", vbYesNo, "File has not been saved")
        If ans = vbYes Then mnuSave_Click
    End If
    ' get an existing banner
    With CommonDialog1
        .Filter = "Banner|*.ban"
        .ShowOpen
        ' if no filename returned then exit sub
        If .FileName = "" Then Exit Sub
        ans = InStrRev(.FileName, ".")
        'set the filename ending with a "."
        If ans = 0 Then
            FileName = .FileName & "."
        Else
            FileName = Left$(.FileName, ans)
        End If
    End With
    ' and open the file
    OpenFile
End Sub

'************************************************************************
' Open a file using FileName
'************************************************************************
Private Sub OpenFile()
    Dim fn As String
    BanPic.Picture = LoadPicture("")
    Open FileName & "ban" For Input As #1
    '========== General Tab =====================
    Input #1, fn: BannerSizes.ListIndex = fn
    If BannerSizes.ListIndex = 0 Then
        Input #1, CustomX, CustomY
    End If
    Input #1, fn: optColor.Value = fn
    Input #1, fn: Display.BackColor = fn
    Input #1, fn: optGraphic.Value = fn
    Input #1, fn: BackgroundPic = fn
    '=========== Text Tab =======================
    Input #1, fn: Text1 = fn
    Input #1, fn: cboFontName.ListIndex = fn
    Input #1, fn: Display.ForeColor = fn
    Input #1, fn: txt1Bold.Value = fn
    Input #1, fn: txt1Italic.Value = fn
    Input #1, fn: txt1Underline.Value = fn
    Input #1, fn: txt1FontSize = fn
    Input #1, fn: Shadow.Value = fn
    Input #1, fn: txtShadow = fn
    Input #1, fn: txtVert = fn
    Input #1, fn: txtHorz = fn
    '=========== Picture Tab ==================
    Input #1, fn: BannerPic = fn
    Input #1, fn: DisplayPic.Value = fn
    Input #1, fn: PictureHeight = fn
    Input #1, fn: PictureWidth = fn
    Input #1, fn: PictureLeft = fn
    Input #1, fn: PictureTop = fn
    ' =========== Border Tab  ============
    Input #1, fn: ShowBorder.Value = fn
    Input #1, fn: BorderWidth = fn
    Input #1, fn: BorderColor = fn
    Input #1, fn: ShadeBorder.Value = fn
    Input #1, fn: LineBorder.Value = fn
    ShadeBorder.Value = Not LineBorder.Value
    
    Close #1
    'load the pictures
    PicBG.Picture = LoadPicture(BackgroundPic)
    BanPic.Picture = LoadPicture(BannerPic)
    Redraw
    Dirty = False
End Sub
'************************************************************************
' Saves the banner file variables
'************************************************************************
Private Sub mnuSave_Click()
    'if the file name is empty then call save as routine
    If FileName = "" Then mnuSaveAs_Click
    If FileName = "" Then Exit Sub
    'open the file and save the values
    Open FileName & "ban" For Output As #1
    '======= General Tab ===========
    Print #1, BannerSizes.ListIndex
    If BannerSizes.ListIndex = 0 Then
        Print #1, CustomX, CustomY
    End If
    Print #1, optColor.Value
    Print #1, Display.BackColor
    Print #1, optGraphic.Value
    Print #1, BackgroundPic
    '=========== Text Tab =============
    Print #1, Text1
    Print #1, cboFontName.ListIndex
    Print #1, Display.ForeColor
    Print #1, txt1Bold.Value
    Print #1, txt1Italic.Value
    Print #1, txt1Underline.Value
    Print #1, txt1FontSize
    Print #1, Shadow.Value
    Print #1, txtShadow
    Print #1, txtVert
    Print #1, txtHorz
    '=========== Picture Tab ===========
    Print #1, BannerPic
    Print #1, DisplayPic.Value
    Print #1, PictureHeight
    Print #1, PictureWidth
    Print #1, PictureLeft
    Print #1, PictureTop
    ' =========== Border Tab  ============
    Print #1, ShowBorder.Value
    Print #1, BorderWidth
    Print #1, BorderColor
    Print #1, ShadeBorder.Value
    Print #1, LineBorder.Value
    'close the file
    Close #1
    'now write the bmp file
    SavePicture Display.Image, FileName & "bmp"
    Dirty = False
End Sub

'************************************************************************
' If a new file with no name is saved or the menu item save file is selected
' then this routine is called.
'************************************************************************
Private Sub mnuSaveAs_Click()
    Dim ans As Integer
    ' get the file name to save to
    With CommonDialog1
        .FileName = ""
        .Filter = "Banner|*.ban"
        .ShowSave
        ' Check if the returned name has something in it
        If .FileName <> "" Then
            ' check if there is a "."
            ans = InStrRev(.FileName, ".")
            If ans = 0 Then
                'if not than add one
                FileName = .FileName & "."
            Else
                ' get rid of anything right of the "."
                FileName = Left$(.FileName, ans)
            End If
        End If
    End With
    ' if a file name was returned then save the file otherwise exit
    If FileName <> "" Then mnuSave_Click
End Sub

'************************************************************************
' Moves the picture to the top left portion of the banner
'************************************************************************
Private Sub MovePicture_Click()
    PictureTop = 0
    PictureLeft = 0
    Dirty = True
End Sub

'************************************************************************
' Sets a line around the banner
'************************************************************************
Private Sub LineBorder_Click()
    'We want to just show a line around the banner
    'the border color button needs to be visible
    Frame5.Visible = True
    'the border shading needs to be invisible
    Frame2.Visible = False
    Dirty = True
    Redraw
End Sub

Private Sub OutlineText_Click()
    Dirty = True
    Redraw
End Sub

'************************************************************************
' Sets the shade around the banner
'************************************************************************
Private Sub ShadeBorder_Click()
    'This shades around the banner to make it look like a button
    'the border color needs to be hidden
    Frame5.Visible = False
    'and the border shad needs to be seen
    Frame2.Visible = True
    Dirty = True
    Redraw
End Sub

'************************************************************************
' Sets if a line or shade is shown on the banner
'************************************************************************
Private Sub ShowBorder_Click()
    Redraw
End Sub

'************************************************************************
' writes each character to the banner as it is typed
'************************************************************************
Private Sub Text1_Change()
    ' sets the max move so the text dosen't move off the banner
    VertMove.Max = Display.ScaleHeight - Display.TextHeight(Text1)
    HorzMove.Max = Display.ScaleWidth - Display.TextWidth(Text1)
    Redraw
    Dirty = True
End Sub

'************************************************************************
' Changes the size of the font
'************************************************************************
Private Sub txt1FontSize_Change()
    Redraw
    Dirty = True
End Sub

'************************************************************************
' if this option is selected then a color background is used in the banner
'************************************************************************
Private Sub optColor_Click()
    Redraw
    Dirty = True
End Sub

'************************************************************************
' if this option is selected than a graphic will be tilled to the background
'************************************************************************
Private Sub optGraphic_Click()
    Redraw
    Dirty = True
End Sub

'************************************************************************
' This option is selected to give the text in the banner a shadow
'************************************************************************
Private Sub Shadow_Click()
    Redraw
    Dirty = True
End Sub

'************************************************************************
' Sets the text in the banner to bold
'************************************************************************
Private Sub txt1Bold_Click()
    Redraw
    Dirty = True
End Sub

'************************************************************************
' sets the text in the banner to italic
'************************************************************************
Private Sub txt1Italic_Click()
    Redraw
    Dirty = True
End Sub

'************************************************************************
' underlines the text in the banner
'************************************************************************
Private Sub txt1Underline_Click()
    Redraw
    Dirty = True
End Sub

'************************************************************************
' the current horizontal position of the text in the banner
'************************************************************************
Private Sub txtHorz_Change()
    Redraw
    Dirty = True
End Sub

'************************************************************************
' how far to offset the shadow from the text in the banner
'************************************************************************
Private Sub txtShadow_Change()
    Redraw
    Dirty = True
End Sub

'************************************************************************
' the current vertical position of the text in the banner
'************************************************************************
Private Sub txtVert_Change()
    Redraw
    Dirty = True
End Sub

'************************************************************************
' Selects the font to use for the text in the banner
'************************************************************************
Private Sub cboFontName_Click()
    Redraw
    Dirty = True
End Sub

'************************************************************************
' option to display a graphic picture in the banner
'************************************************************************
Private Sub DisplayPic_Click()
    Redraw
    Dirty = True
End Sub

'************************************************************************
' Get the file for the picture in the banner
'************************************************************************
Private Sub cmdBrowsePic_Click()
    With CommonDialog1
        .Filter = "bmp,jpg,gif,ico,wmf|*.bmp;*.jpg;*.gif;*.ico;*.wmf"
        .ShowOpen
        If CommonDialog1.FileName = "" Then Exit Sub
        ' Load the picture
        BanPic.Picture = LoadPicture(.FileName)
        PictureHeight = BanPic.ScaleHeight
        PictureWidth = BanPic.Width * PictureHeight / BanPic.Height
        ' Save the picture filename
        BannerPic = .FileName
    End With
    DisplayPic.Enabled = True
    Redraw
    Dirty = True
End Sub

'************************************************************************
' Makes the graphic picture larger
'************************************************************************
Private Sub ResizeUp_Click()
    PictureHeight = PictureHeight + Val(Text2)
    Redraw
    Dirty = True
End Sub

'************************************************************************
' Makes the graphic picture smaller
'************************************************************************
Private Sub ResizeDown_Click()
    PictureHeight = PictureHeight - Val(Text2)
    Redraw
    Dirty = True
End Sub

'************************************************************************
' Moves the graphic picture down
'************************************************************************
Private Sub MoveDown_Click()
    PictureTop = PictureTop + Val(Text3)
    Redraw
    Dirty = True
End Sub

'************************************************************************
' Moves the graphic picture up
'************************************************************************
Private Sub MoveUP_Click()
    PictureTop = PictureTop - Val(Text3)
    Redraw
    Dirty = True
End Sub

'************************************************************************
' Moves the graphic picture left
'************************************************************************
Private Sub MoveLeft_Click()
    PictureLeft = PictureLeft - Val(Text3)
    Redraw
    Dirty = True
End Sub

'************************************************************************
' Moves the graphic picture right
'************************************************************************
Private Sub MoveRight_Click()
    PictureLeft = PictureLeft + Val(Text3)
    Redraw
    Dirty = True
End Sub


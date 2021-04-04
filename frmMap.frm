VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMap 
   AutoRedraw      =   -1  'True
   Caption         =   "Dungeon Map Maker"
   ClientHeight    =   10005
   ClientLeft      =   405
   ClientTop       =   495
   ClientWidth     =   13590
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMap.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMap.frx":306A
   ScaleHeight     =   667
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   906
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1140
      TabIndex        =   67
      Top             =   7920
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   66
      Top             =   7920
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdText 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   65
      ToolTipText     =   "Add Text to Map"
      Top             =   4620
      Width           =   375
   End
   Begin VB.CommandButton cmdMerge 
      Caption         =   "Merge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   58
      Top             =   9540
      Width           =   735
   End
   Begin VB.PictureBox picCurrent 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      Picture         =   "frmMap.frx":3818
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   64
      ToolTipText     =   "Archway North & East"
      Top             =   5280
      Width           =   375
   End
   Begin VB.PictureBox picProduct 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1560
      Picture         =   "frmMap.frx":3FC6
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   61
      ToolTipText     =   "Archway North & East"
      Top             =   9120
      Width           =   375
   End
   Begin VB.PictureBox picReversedMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   120
      Picture         =   "frmMap.frx":4774
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   60
      ToolTipText     =   "Archway North & East"
      Top             =   6180
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox picBackgd 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   720
      Picture         =   "frmMap.frx":4F22
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   59
      ToolTipText     =   "Archway North & East"
      Top             =   9120
      Width           =   375
   End
   Begin VB.PictureBox picRevMaskedFore 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   120
      Picture         =   "frmMap.frx":56D0
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   57
      ToolTipText     =   "Archway North & East"
      Top             =   6660
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   720
      Picture         =   "frmMap.frx":5E7E
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   56
      ToolTipText     =   "Archway North & East"
      Top             =   6180
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox picFgd 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   60
      Picture         =   "frmMap.frx":662C
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   55
      ToolTipText     =   "Archway North & East"
      Top             =   9120
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   53
      Left            =   420
      Picture         =   "frmMap.frx":6DDA
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   54
      ToolTipText     =   "Sarcophagus Open"
      Top             =   4620
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   51
      Left            =   1680
      Picture         =   "frmMap.frx":708C
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   53
      ToolTipText     =   "Sarcophagus Open"
      Top             =   4200
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   38
      Left            =   1260
      Picture         =   "frmMap.frx":733E
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   52
      ToolTipText     =   "Catapult"
      Top             =   2940
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   29
      Left            =   1680
      Picture         =   "frmMap.frx":7AEC
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   51
      ToolTipText     =   "Well"
      Top             =   2100
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   1680
      Picture         =   "frmMap.frx":829A
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   50
      ToolTipText     =   "Trap"
      Top             =   420
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   46
      Left            =   840
      Picture         =   "frmMap.frx":8A48
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   49
      ToolTipText     =   "Alter"
      Top             =   3780
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   44
      Left            =   0
      Picture         =   "frmMap.frx":8CC5
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   48
      ToolTipText     =   "Treasure"
      Top             =   3780
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   39
      Left            =   1680
      Picture         =   "frmMap.frx":8FCF
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   47
      ToolTipText     =   "Treasure Chest"
      Top             =   2940
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   40
      Left            =   0
      Picture         =   "frmMap.frx":977D
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   46
      ToolTipText     =   "Teleport Area"
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1680
      Picture         =   "frmMap.frx":9981
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   45
      ToolTipText     =   "Open Pit"
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   1680
      Picture         =   "frmMap.frx":A12F
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   44
      ToolTipText     =   "Pillar"
      Top             =   1680
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   1680
      Picture         =   "frmMap.frx":A8DD
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   43
      ToolTipText     =   "Closed Pit"
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   1680
      Picture         =   "frmMap.frx":B08B
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   42
      ToolTipText     =   "Secret Trap Door"
      Top             =   1260
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   1260
      Picture         =   "frmMap.frx":B37B
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   41
      ToolTipText     =   "Pool/Fountian"
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   1260
      Picture         =   "frmMap.frx":B635
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   40
      ToolTipText     =   "Water"
      Top             =   1260
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   35
      Left            =   0
      Picture         =   "frmMap.frx":BDE3
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   39
      ToolTipText     =   "Cave Openning"
      Top             =   2940
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   47
      Left            =   0
      Picture         =   "frmMap.frx":C591
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   38
      ToolTipText     =   "Sarcophagus Open"
      Top             =   4200
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   0
      Picture         =   "frmMap.frx":C843
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   37
      ToolTipText     =   "Cave intersection"
      Top             =   1680
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   49
      Left            =   840
      Picture         =   "frmMap.frx":CFF1
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   36
      ToolTipText     =   "Curtian East/West"
      Top             =   4200
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   52
      Left            =   0
      Picture         =   "frmMap.frx":D250
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   35
      ToolTipText     =   "Curtian North/South"
      Top             =   4620
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   1260
      Picture         =   "frmMap.frx":D4B5
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   34
      ToolTipText     =   "Dirt Ground"
      Top             =   1680
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   30
      Left            =   0
      Picture         =   "frmMap.frx":DC63
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   33
      ToolTipText     =   "Cave Wall"
      Top             =   2520
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   50
      Left            =   1260
      Picture         =   "frmMap.frx":E411
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   32
      ToolTipText     =   "Wild Magic Area"
      Top             =   4200
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   34
      Left            =   1680
      Picture         =   "frmMap.frx":E628
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   31
      ToolTipText     =   "Spiral Stairs"
      Top             =   2520
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   42
      Left            =   840
      Picture         =   "frmMap.frx":EDD6
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   30
      ToolTipText     =   "Thick Wall West"
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   33
      Left            =   1260
      Picture         =   "frmMap.frx":EFB3
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   29
      ToolTipText     =   "Bastilla"
      Top             =   2520
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   37
      Left            =   840
      Picture         =   "frmMap.frx":F761
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   28
      ToolTipText     =   "Castle Wall Corner"
      Top             =   2940
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   31
      Left            =   420
      Picture         =   "frmMap.frx":FF0F
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   27
      ToolTipText     =   "Large Open Doorway Left"
      Top             =   2520
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   7
      Left            =   840
      Picture         =   "frmMap.frx":106BD
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   26
      ToolTipText     =   "Hall Way"
      Top             =   420
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   840
      Picture         =   "frmMap.frx":10E6B
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   25
      ToolTipText     =   "Large Doorway Right"
      Top             =   2100
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   840
      Picture         =   "frmMap.frx":11619
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   24
      ToolTipText     =   "Open Hallway Door"
      Top             =   1680
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   420
      Picture         =   "frmMap.frx":11DC7
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   23
      ToolTipText     =   "Room Corner"
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   420
      Picture         =   "frmMap.frx":12575
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   22
      ToolTipText     =   "Large Open Doorway Right"
      Top             =   2100
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   0
      Picture         =   "frmMap.frx":12D23
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   21
      ToolTipText     =   "Black"
      Top             =   420
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   32
      Left            =   840
      Picture         =   "frmMap.frx":134D1
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   20
      ToolTipText     =   "Large Doorway Left"
      Top             =   2520
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   840
      Picture         =   "frmMap.frx":13C7F
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   19
      ToolTipText     =   "Secret Hallway Door"
      Top             =   1260
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   0
      Picture         =   "frmMap.frx":1442D
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   18
      ToolTipText     =   "Bottom Door"
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   420
      Picture         =   "frmMap.frx":14BDB
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   17
      ToolTipText     =   "Room Wall"
      Top             =   420
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   28
      Left            =   1260
      Picture         =   "frmMap.frx":15389
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   16
      ToolTipText     =   "Bars"
      Top             =   2100
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   420
      Picture         =   "frmMap.frx":15B37
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   15
      ToolTipText     =   "Secret Passage West"
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   45
      Left            =   420
      Picture         =   "frmMap.frx":162E5
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   14
      ToolTipText     =   "East/West Stairs"
      Top             =   3780
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1260
      Picture         =   "frmMap.frx":1651A
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   13
      ToolTipText     =   "Stairs Down"
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1260
      Picture         =   "frmMap.frx":16CC8
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   12
      ToolTipText     =   "Stairs Up"
      Top             =   420
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   36
      Left            =   420
      Picture         =   "frmMap.frx":17476
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   11
      ToolTipText     =   "Castle Wall"
      Top             =   2940
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   43
      Left            =   1320
      Picture         =   "frmMap.frx":17C24
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   10
      ToolTipText     =   "Large Tower Wall"
      Top             =   3360
      Width           =   750
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   420
      Picture         =   "frmMap.frx":19A16
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   9
      ToolTipText     =   "Secret Door"
      Top             =   1260
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   420
      Picture         =   "frmMap.frx":1A1C4
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   8
      ToolTipText     =   "Open Doorway"
      Top             =   1680
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   840
      Picture         =   "frmMap.frx":1A972
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   7
      ToolTipText     =   "Dead-end Hall"
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   0
      Picture         =   "frmMap.frx":1B120
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   6
      ToolTipText     =   "Cave Tunnel"
      Top             =   2100
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   0
      Picture         =   "frmMap.frx":1B8CE
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   5
      ToolTipText     =   "Cave Corner"
      Top             =   1260
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   41
      Left            =   420
      Picture         =   "frmMap.frx":1C07C
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   4
      ToolTipText     =   "Round Tower wall"
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   48
      Left            =   420
      Picture         =   "frmMap.frx":1C82A
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   3
      ToolTipText     =   "Archway East/West"
      Top             =   4200
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   0
      Picture         =   "frmMap.frx":1CB64
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   2
      ToolTipText     =   "Empty"
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   840
      Picture         =   "frmMap.frx":1D312
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   1
      ToolTipText     =   "Archway North & East"
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   10785
      Left            =   2160
      ScaleHeight     =   715
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   870
      TabIndex        =   0
      Top             =   0
      Width           =   13110
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6540
      Top             =   4740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picmap2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10785
      Left            =   2160
      ScaleHeight     =   715
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   870
      TabIndex        =   68
      Top             =   -600
      Visible         =   0   'False
      Width           =   13110
   End
   Begin VB.Label Label2 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   63
      Top             =   9180
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   62
      Top             =   9180
      Width           =   255
   End
   Begin VB.Menu mnuopt 
      Caption         =   "&Options"
      Begin VB.Menu mnunew 
         Caption         =   "&New"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuopen 
         Caption         =   "&Open"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Save"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lXText As Long
Dim lYText As Long
Dim bTextMode As Boolean

Dim lXPos As Long
    'this carries the X ( or left-right for all us dislect information which the tiles
    'will be BitBlted on
Dim lYPos As Long
    'this carries the Y information which the tiles
    'will be BitBlted on
Dim sAddText As String
Private Sub cmdAdd_Click()
    bTextMode = False
    cmdText.Enabled = True
    Me.cmdAdd.Visible = False
    Me.cmdCancel.Visible = False
    sAddText = ""
    lXText = 0
End Sub
Private Sub cmdCancel_Click()
    bTextMode = False
    cmdText.Enabled = True
    BitBlt Me.picMap.hdc, 0, 0, picmap2.Width, picmap2.Height, picmap2.hdc, 0, 0, vbSrcCopy
    picMap.Refresh
    cmdCancel.Visible = False
    cmdAdd.Visible = False
    sAddText = ""
End Sub

Private Sub cmdRotate_Click()
    Dim picTemp As PictureBox
    Dim RET As Long
    
    Set picTemp = picCurrent
   ' RET = BitBlt(picRotate.hdc, 0, 0, picCurrent.Height, picCurrent.Width, picCurrent.hdc, 0, 0, SRCCOPY)
    RotatePicture picCurrent, picTemp
    RET = BitBlt(picCurrent.hdc, 0, 0, picTemp.Height, picTemp.Width, picTemp.hdc, 0, 0, SRCCOPY)
    picCurrent.Refresh
    Set picTemp = Nothing
End Sub
Private Sub cmdMerge_Click()

    BitBlt picMask.hdc, 0, 0, picFgd.Height, picFgd.Width, picFgd.hdc, 0, 0, vbSrcCopy
    picMask.Refresh
    
     ' Do masking
    CreateMask picMask, vbBlack
    
    ' Background picBackgd can readily be copied onto picProduct
    BitBlt picProduct.hdc, 0, 0, picBackgd.Height, picBackgd.Width, picBackgd.hdc, 0, 0, vbSrcCopy
    picProduct.Picture = picProduct.Image
    
    ' Copy the mask onto the picProduct using the vbMergePaint opcode
    ' to erase pixels corresponding to black parts of the mask.
    BitBlt picProduct.hdc, 0, 0, picMask.Height, picMask.Width, picMask.hdc, 0, 0, vbMergePaint
    picProduct.Picture = picProduct.Image

    CreateReverseMaskedFgd
    
    ' Copy the reverse masked Fgd image onto the masked background
    BitBlt picProduct.hdc, 0, 0, picRevMaskedFore.Height, picRevMaskedFore.Width, picRevMaskedFore.hdc, _
          0, 0, vbSrcAnd
    picProduct.Picture = picProduct.Image
    
End Sub


Private Sub cmdText_Click()
    bTextMode = True
    lXText = 0
    BitBlt picmap2.hdc, 0, 0, picMap.Width, picMap.Height, picMap.hdc, 0, 0, vbSrcCopy
    cmdAdd.Visible = True
    cmdCancel.Visible = True
    cmdText.Enabled = False
End Sub

Private Sub mnuexit_Click()
    Unload Me
End Sub
Private Sub mnunew_Click()
    picMap.Picture = LoadPicture("")
End Sub
Private Sub mnuopen_Click()
    On Error GoTo openerr
    CommonDialog1.Filter = "BitMap Files (*.bmp)| *.bmp"
    CommonDialog1.ShowOpen
    picMap.Picture = LoadPicture(CommonDialog1.FileName)
    frmMap.Caption = "Dungeon Map Maker.. " & CommonDialog1.FileName
    Exit Sub
openerr:
    Exit Sub
End Sub

Private Sub mnusave_Click()
    Dim Mypicture As String
    Randomize Timer
    Mypicture = App.Path & "\Map" & CStr(Int(Rnd * 500)) & ".bmp"
    If Dir(Mypicture) <> "" Then
        mnusave_Click
    End If
    SavePicture picMap.Image, Mypicture
    frmMap.Caption = "Dungeon Map Maker "
End Sub

Private Sub picBackgd_Click()
    BitBlt picBackgd.hdc, 0, 0, picCurrent.Height, picCurrent.Width, picCurrent.hdc, 0, 0, SRCCOPY
    picBackgd.Refresh
End Sub

Private Sub picCurrent_Click()
    Dim picTemp As PictureBox
    Set picTemp = picCurrent
    RotatePicture picCurrent, picTemp
    BitBlt picCurrent.hdc, 0, 0, picTemp.Height, picTemp.Width, picTemp.hdc, 0, 0, SRCCOPY
    picCurrent.Refresh
    Set picTemp = Nothing
End Sub

Private Sub picFgd_Click()
    BitBlt picFgd.hdc, 0, 0, picCurrent.Height, picCurrent.Width, picCurrent.hdc, 0, 0, SRCCOPY
    picFgd.Refresh
End Sub

Private Sub picMap_KeyPress(KeyAscii As Integer)

    If Len(sAddText) > 0 And KeyAscii = 8 Then
        sAddText = Left$(sAddText, Len(sAddText) - 1)
        BitBlt Me.picMap.hdc, 0, 0, picmap2.Width, picmap2.Height, picmap2.hdc, 0, 0, vbSrcCopy
    Else
        If KeyAscii <> 8 Then
            sAddText = sAddText & Chr(KeyAscii)
        End If
    End If
    TextTrans
End Sub

Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim TempDC As Long
    Dim Temp As String
    Dim MyLoc As RECT
    
    If bTextMode Then
        If lXText = 0 Then
            lXText = x
            lYText = y
        End If
    Else
        If Button = 2 Then
            cmdRotate_Click
        End If
        Call PaintTile(lXPos, lYPos)
    End If
End Sub

Private Sub TextTrans()
    Dim TempDC As Long
    Dim MyLoc As RECT
    Dim BlOC As RECT
    MyLoc.Left = lXText
    MyLoc.Top = lYText
    MyLoc.Right = MyLoc.Left + Me.TextWidth(sAddText) * 5
    MyLoc.Bottom = MyLoc.Top + Me.TextHeight(sAddText)
    DrawText picMap.hdc, " " & sAddText & " ", Len(sAddText) + 2, MyLoc, DT_EDITCONTROL
    picMap.Refresh
End Sub


Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lXPos = (x \ 25) * 25
    'lXPos is the current mouse's X value, but divided
    'by the TileSize, returning INTEGER only, that means
    'no decimals! That is what the sign "\" is for,
    'instead of "/" that you usually see
    
    lYPos = (y \ 25) * 25
    'same as above, replace X with Y
    If bTextMode Then
        Exit Sub
    End If

    If Button = 1 Then
        'if the user has the mouse button down, then call the sub PaintTile
        PaintTile lXPos, lYPos
    Else
        Exit Sub
    End If
End Sub
Private Sub PaintTile(ByVal SrcX As Long, ByVal SrcY As Long)
    'this BitBlts to the picMap, with the information given when this sub was called
    BitBlt picMap.hdc, SrcX, SrcY, picCurrent.Height, picCurrent.Width, picCurrent.hdc, 0, 0, SRCCOPY
    picMap.Refresh
End Sub

Private Sub picProduct_Click()
    BitBlt picCurrent.hdc, 0, 0, picProduct.Height, picProduct.Width, picProduct.hdc, 0, 0, SRCCOPY
    picCurrent.Picture = picProduct.Picture
    picCurrent.Refresh
End Sub

Private Sub picTile_Click(Index As Integer)

    picCurrent.Height = picTile(Index).Height
    picCurrent.Width = picTile(Index).Width
    BitBlt picCurrent.hdc, 0, 0, picTile(Index).Height, picTile(Index).Width, picTile(Index).hdc, 0, 0, SRCCOPY
    picCurrent.Picture = picTile(Index).Picture
    picCurrent.Refresh
End Sub

Sub CreateMask(inPic As PictureBox, inColorToUse)
    'On Error Resume Next
    Dim mTranspColor As Long
    Dim iWidth As Integer
    Dim iHeight As Integer
    
    mTranspColor = inPic.Point(0, 0)
        ' See if existing background is fully covered by
        ' some foreground color which is to serve as
        ' background visually. We are to use image of
        ' picBackgd as the background.
    If mTranspColor <> inColorToUse Then
        For iHeight = 0 To inPic.Height + 1
            For iWidth = 0 To inPic.Width + 1
                If inPic.Point(iHeight, iWidth) = mTranspColor Then
                    inPic.PSet (iHeight, iWidth), vbWhite
                End If
            Next iWidth
            DoEvents
        Next iHeight
    End If
    
    For iHeight = 0 To inPic.Height + 1
        For iWidth = 0 To inPic.Width + 1
            If inPic.Point(iHeight, iWidth) <> vbWhite Then
                inPic.PSet (iHeight, iWidth), inColorToUse
            End If
        Next iWidth
        DoEvents
    Next iHeight
End Sub

Private Sub CreateReverseMaskedFgd()
    ' Make a reversed mask.
    BitBlt picReversedMask.hdc, 0, 0, picMask.Height, picMask.Width, picMask.hdc, 0, 0, vbNotSrcCopy
    picReversedMask.Picture = picReversedMask.Image

    ' Copy picFgd to picRevMaskedFore
    BitBlt picRevMaskedFore.hdc, 0, 0, picFgd.Height, picFgd.Width, picFgd.hdc, _
        0, 0, vbSrcCopy
    picRevMaskedFore.Picture = picRevMaskedFore.Image

    ' Copy the earlier reversed mask onto the picRevserseMaskedFgd
    ' using vbMergePaint opcode to erase part of the foreground
    ' which corresponds to the black parts of that reversed mask.
    BitBlt picRevMaskedFore.hdc, 0, 0, picReversedMask.Height, picReversedMask.Width, picReversedMask.hdc, _
           0, 0, vbMergePaint
    picRevMaskedFore.Picture = picRevMaskedFore.Image
End Sub


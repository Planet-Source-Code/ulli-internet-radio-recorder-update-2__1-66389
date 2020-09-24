VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E8E8FB&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   5100
   ClientLeft      =   2055
   ClientTop       =   2370
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E8E8FB&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   340
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox ckStick 
      Height          =   195
      Left            =   5835
      TabIndex        =   11
      ToolTipText     =   "Discard all Titles"
      Top             =   1905
      Width           =   210
   End
   Begin VB.CommandButton btAbout 
      BackColor       =   &H00C0FFFF&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5295
      Style           =   1  'Grafisch
      TabIndex        =   13
      ToolTipText     =   "Show About box"
      Top             =   2250
      Width           =   810
   End
   Begin VB.Timer tmrOff 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4874
      Top             =   4635
   End
   Begin VB.Timer tmrAnim 
      Interval        =   100
      Left            =   4447
      Top             =   4635
   End
   Begin VB.CommandButton btRun 
      BackColor       =   &H00A0FFA0&
      Caption         =   "Start Receiving"
      Enabled         =   0   'False
      Height          =   420
      Left            =   4440
      Style           =   1  'Grafisch
      TabIndex        =   9
      Top             =   1365
      Width           =   1665
   End
   Begin VB.CommandButton btRestart 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Restart"
      Height          =   405
      Left            =   4440
      Style           =   1  'Grafisch
      TabIndex        =   12
      ToolTipText     =   "Restart Playback from Beginning"
      Top             =   2250
      Width           =   825
   End
   Begin VB.PictureBox picIcon 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      Height          =   480
      Left            =   3960
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   49
      Top             =   4635
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CheckBox ckUseTitleList 
      Alignment       =   1  'Rechts ausgerichtet
      BackColor       =   &H00FFF0EA&
      Caption         =   "Use Title List -->"
      ForeColor       =   &H00707070&
      Height          =   225
      Left            =   705
      TabIndex        =   14
      ToolTipText     =   "Rightclick for List"
      Top             =   4275
      Width           =   1500
   End
   Begin VB.CommandButton btSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   6.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5850
      TabIndex        =   6
      ToolTipText     =   "Browse"
      Top             =   1005
      Width           =   255
   End
   Begin VB.CommandButton btEdit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5850
      Style           =   1  'Grafisch
      TabIndex        =   15
      ToolTipText     =   "Edit Station List"
      Top             =   345
      Width           =   255
   End
   Begin VB.Frame fr 
      BackColor       =   &H00FF90FF&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H00000000&
      Height          =   405
      Index           =   3
      Left            =   5700
      TabIndex        =   42
      Top             =   1380
      Width           =   405
      Begin VB.PictureBox picAnim 
         BorderStyle     =   0  'Kein
         Height          =   375
         Index           =   1
         Left            =   15
         Picture         =   "frmMain.frx":1194
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   15
         Width           =   375
      End
      Begin VB.PictureBox picAnim 
         BorderStyle     =   0  'Kein
         Height          =   360
         Index           =   0
         Left            =   15
         Picture         =   "frmMain.frx":1896
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   15
         Width           =   360
      End
      Begin VB.PictureBox picAnim 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'Kein
         Height          =   360
         Index           =   2
         Left            =   15
         Picture         =   "frmMain.frx":1FF8
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   15
         Width           =   360
      End
   End
   Begin VB.Frame fr 
      BackColor       =   &H00E8FBFB&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00707070&
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   35
      ToolTipText     =   "Status Area"
      Top             =   2175
      Width           =   4095
      Begin VB.Label lbOut 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Recording Limit Reached"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   135
         TabIndex        =   48
         Top             =   195
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lbInterrupted 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "No Net"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   2055
         TabIndex        =   47
         Top             =   195
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbWillBe 
         Alignment       =   1  'Rechts
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Will be"
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   2760
         TabIndex        =   39
         Top             =   195
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lbWriting 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Recording"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   135
         TabIndex        =   38
         Top             =   195
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lbBuffering 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Buffering"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   1110
         TabIndex        =   37
         Top             =   195
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label lbDiscard 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Discarded"
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   3255
         TabIndex        =   36
         Top             =   195
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin VB.Frame fr 
      BackColor       =   &H00E8E8FB&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   105
      TabIndex        =   33
      Top             =   690
      Width           =   6000
      Begin VB.OptionButton optAllInOneFile 
         BackColor       =   &H00E8E8FB&
         Caption         =   "Fixed Output File:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505070&
         Height          =   195
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "All Titles into one File"
         Top             =   30
         Width           =   1755
      End
      Begin VB.OptionButton optSeparateFiles 
         BackColor       =   &H00E8E8FB&
         Caption         =   "Separate Output File for each Title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505070&
         Height          =   195
         Left            =   2850
         TabIndex        =   4
         ToolTipText     =   "Separate File for each Title"
         Top             =   30
         Value           =   -1  'True
         Width           =   3210
      End
   End
   Begin VB.ComboBox cboStations 
      BackColor       =   &H00E0F8FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      ItemData        =   "frmMain.frx":26FA
      Left            =   90
      List            =   "frmMain.frx":26FC
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Station or URL"
      Top             =   330
      Width           =   5790
   End
   Begin MSComctlLib.ProgressBar Bar 
      Height          =   210
      Left            =   135
      TabIndex        =   19
      ToolTipText     =   "Progess"
      Top             =   2835
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox txtSave 
      BackColor       =   &H00E0F8FF&
      Enabled         =   0   'False
      ForeColor       =   &H00400040&
      Height          =   285
      Left            =   105
      MaxLength       =   160
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Output File Name"
      Top             =   975
      Width           =   6030
   End
   Begin VB.Frame fr 
      BackColor       =   &H00E8E8FB&
      Caption         =   "Recording Limit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00707070&
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "Recording Limit Area"
      Top             =   1305
      Width           =   4095
      Begin VB.TextBox txtBreak 
         Alignment       =   1  'Rechts
         BackColor       =   &H00E0F8FF&
         Height          =   285
         Index           =   1
         Left            =   2385
         MaxLength       =   4
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "100"
         ToolTipText     =   "Megabytes"
         Top             =   270
         Width           =   495
      End
      Begin VB.TextBox txtBreak 
         Alignment       =   1  'Rechts
         BackColor       =   &H00E0F8FF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   495
         MaxLength       =   3
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "60"
         ToolTipText     =   "Minutes"
         Top             =   270
         Width           =   495
      End
      Begin VB.OptionButton optBreak 
         BackColor       =   &H00E8E8FB&
         Caption         =   "              Megabytes"
         ForeColor       =   &H00505070&
         Height          =   225
         Index           =   1
         Left            =   2100
         TabIndex        =   8
         ToolTipText     =   "Volume Limit"
         Top             =   315
         Width           =   1755
      End
      Begin VB.OptionButton optBreak 
         BackColor       =   &H00E8E8FB&
         Caption         =   "              Minutes"
         ForeColor       =   &H00505070&
         Height          =   225
         Index           =   0
         Left            =   195
         TabIndex        =   7
         ToolTipText     =   "Time Limit"
         Top             =   300
         Value           =   -1  'True
         Width           =   1515
      End
      Begin VB.Line ln 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   1920
         X2              =   1920
         Y1              =   105
         Y2              =   705
      End
      Begin VB.Line ln 
         BorderColor     =   &H00A0A0A0&
         Index           =   0
         X1              =   1905
         X2              =   1905
         Y1              =   90
         Y2              =   705
      End
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   5790
      Top             =   4635
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5301
      Top             =   4635
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox ckKill 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Discard Current   Title"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   4440
      Style           =   1  'Grafisch
      TabIndex        =   10
      ToolTipText     =   "Discard current title"
      Top             =   1800
      Width           =   1665
   End
   Begin VB.Label lbVersion 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C080C0&
      Height          =   135
      Left            =   5775
      TabIndex        =   46
      Top             =   -15
      Width           =   465
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00707070&
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   41
      Top             =   4275
      Width           =   420
   End
   Begin VB.Label lbCurrTime 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0:00"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   1650
      TabIndex        =   34
      Top             =   4725
      Width           =   330
   End
   Begin VB.Label lbCurrTime 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0:00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1665
      TabIndex        =   40
      Top             =   4740
      Width           =   330
   End
   Begin VB.Line ln 
      BorderColor     =   &H00FFB0FF&
      Index           =   5
      X1              =   288
      X2              =   418
      Y1              =   181
      Y2              =   181
   End
   Begin VB.Line ln 
      BorderColor     =   &H00FFB0FF&
      Index           =   4
      X1              =   287
      X2              =   287
      Y1              =   144
      Y2              =   182
   End
   Begin VB.Line ln 
      BorderColor     =   &H00FFB0FF&
      Index           =   2
      X1              =   0
      X2              =   288
      Y1              =   143
      Y2              =   143
   End
   Begin VB.Line ln 
      BorderColor     =   &H00FFC0C0&
      Index           =   3
      X1              =   0
      X2              =   417
      Y1              =   252
      Y2              =   252
   End
   Begin VB.Label lbStream 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[Waiting  for Station  Info]"
      ForeColor       =   &H00FF4040&
      Height          =   195
      Left            =   135
      MouseIcon       =   "frmMain.frx":26FE
      TabIndex        =   30
      Top             =   4005
      Width           =   1905
   End
   Begin VB.Shape shp 
      BorderColor     =   &H00E8FBFB&
      FillColor       =   &H00E8FBFB&
      FillStyle       =   0  'Ausgefüllt
      Height          =   585
      Index           =   3
      Left            =   0
      Top             =   2145
      Width           =   4320
   End
   Begin VB.Shape shp 
      BorderColor     =   &H0000C0C0&
      FillColor       =   &H00DDDDFF&
      Height          =   240
      Index           =   0
      Left            =   120
      Top             =   2820
      Width           =   5985
   End
   Begin VB.Label lbSong 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[Waiting for Title]"
      ForeColor       =   &H00D05030&
      Height          =   195
      Left            =   135
      MouseIcon       =   "frmMain.frx":2A08
      TabIndex        =   32
      Top             =   4455
      Width           =   1260
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Station Info:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00707070&
      Height          =   195
      Index           =   7
      Left            =   135
      TabIndex        =   31
      Top             =   3810
      Width           =   1050
   End
   Begin VB.Label lbGenre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[Unknown]"
      ForeColor       =   &H00008080&
      Height          =   195
      Left            =   3495
      TabIndex        =   29
      Top             =   3195
      Width           =   780
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Stream Genre:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00707070&
      Height          =   195
      Index           =   3
      Left            =   2100
      TabIndex        =   27
      Top             =   3195
      Width           =   1230
   End
   Begin VB.Label lbPackets 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00008080&
      Height          =   195
      Left            =   1335
      TabIndex        =   28
      Top             =   3195
      Width           =   90
   End
   Begin VB.Label lbBitrate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[Unknown]"
      ForeColor       =   &H00008080&
      Height          =   195
      Left            =   5355
      TabIndex        =   26
      Top             =   3450
      Width           =   780
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Audio Bitrate:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00707070&
      Height          =   195
      Index           =   6
      Left            =   4095
      TabIndex        =   25
      Top             =   3450
      Width           =   1155
   End
   Begin VB.Label lbTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0:00"
      ForeColor       =   &H00008080&
      Height          =   195
      Left            =   3495
      TabIndex        =   24
      Top             =   3450
      Width           =   330
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Audio Duration:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00707070&
      Height          =   195
      Index           =   5
      Left            =   2100
      TabIndex        =   22
      Top             =   3450
      Width           =   1305
   End
   Begin VB.Label lbWritten 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 kB"
      ForeColor       =   &H00008080&
      Height          =   195
      Left            =   1335
      TabIndex        =   23
      Top             =   3450
      Width           =   300
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Data stored:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00707070&
      Height          =   195
      Index           =   4
      Left            =   135
      TabIndex        =   21
      Top             =   3450
      Width           =   1050
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Packets rcvd:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00707070&
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   20
      Top             =   3195
      Width           =   1140
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Station:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00707070&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   660
   End
   Begin VB.Shape shp 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFF0EA&
      FillStyle       =   0  'Ausgefüllt
      Height          =   870
      Index           =   2
      Left            =   0
      Top             =   3795
      Width           =   6270
   End
   Begin VB.Shape shp 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E8FBFB&
      FillStyle       =   0  'Ausgefüllt
      Height          =   1125
      Index           =   1
      Left            =   0
      Top             =   2730
      Width           =   6270
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmPlayer 
      Height          =   675
      Left            =   -840
      TabIndex        =   1
      Top             =   4425
      Width           =   7155
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "mini"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   12621
      _cy             =   1191
   End
   Begin VB.Menu mnuTray 
      Caption         =   "TrayMenu"
      NegotiatePosition=   2  'Mitte
      Visible         =   0   'False
      Begin VB.Menu mnuDefault 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnusep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStartStop 
         Caption         =   "Start Receiving"
      End
      Begin VB.Menu mnuKill 
         Caption         =   "Discard Current Title"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRestart 
         Caption         =   "Restart Playback"
      End
      Begin VB.Menu mnuBalloons 
         Caption         =   "Balloons"
         Begin VB.Menu mnuSuppBalloons 
            Caption         =   "Suppress"
         End
         Begin VB.Menu mnuEnaBallons 
            Caption         =   "Enable"
         End
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSendMail 
         Caption         =   "Send Mail to Author"
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDummy 
         Caption         =   "Hide Menu"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Send Mail
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long

'About box
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

'process check for notepad
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'misc
Private Declare Function FileExists Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Enum ApiConsts
    WM_SETFOCUS = 7
    WM_CLOSE = &H10
    HWND_TOPMOST = -1
    HWND_NOTOPMOST = -2
    PROCESS_ALL_ACCESS = &H1F0FFF
    SWP_NOSIZE = 1
    SWP_NOMOVE = 2
    SWP_NONOTHING = SWP_NOSIZE Or SWP_NOMOVE
    SW_SHOWNORMAL = 1
    SE_NO_ERROR = 33 'Values below 33 are error returns
End Enum
#If False Then
Private WM_SETFOCUS, WM_CLOSE, HWND_TOPMOST, HWND_NOTOPMOST, PROCESS_ALL_ACCESS, SWP_NOSIZE, SWP_NOMOVE, SWP_NONOTHING, SW_SHOWNORMAL, SE_NO_ERROR
#End If

Private Enum PartId
    Label = 0
    Contents = 1
End Enum
#If False Then
Private Label, Contents
#End If

Private WithEvents Systray          As clsSystray
Attribute Systray.VB_VarHelpID = -1
Private EmergencyUnhook             As clsUnhook 'unhooking thru a class on destruction; this has the
'                                                'advantage that it will also unhook un unexpected errors
Private hFile                       As Long
Private AnimPosn                    As Long
Private AnimCounter                 As Long 'counts the anim ticks and stops animation unless reset by a packet
Private MetaPacketSize              As Long
Private BitrateInKBits              As Long
Private BitrateInBits               As Long
Private ByterateInBytes             As Long
Private NumPacketsReceived          As Long
Private NumBytesReceived            As Long
Private MaxBufferSize               As Long
Private AvgPacketSize               As Long
Private MaxPacketSize               As Long
Private NumBytesAppended            As Long 'the total number of bytes transfered from input to output
Private NumBytesWritten             As Long 'the number of bytes written during a title
Private NumBlocksWritten            As Long 'the number of disc blocks written during a title
Private NumSessionBlocksWritten     As Long 'the total number of blocks written
Private NumBytesAppendedForTitle    As Long 'the number of bytes transferred during a title

Private ICYHeaderReceived           As Boolean
Private Running                     As Boolean
Private ByCommandline               As Boolean
Private Agreed                      As Boolean
Private BalloonsEnabled             As Boolean

Private Stations()                  As String
Private Titles()                    As String
Private CurrStation                 As String
Private OutputFilename              As String
Private Packet                      As String
Private InputBuffer                 As String
Private Metablock                   As String
Private WriteBuffer                 As String

Private WaitingStation              As String
Private WaitingTitle                As String

Private Const MaxWriteBufferSize    As Long = 65536 '64k
Private Const AnimSpeed             As Long = 100
Private Const AnimTimeout           As Long = 101
Private Const EndHeader             As String = vbCrLf & vbCrLf
Private Const fnNetStations         As String = "NetStations.txt"
Private Const fnTitles              As String = "Titles.txt"
Private Const MusicDir              As String = "\mp3\"
Private Const MP3                   As String = ".mp3"
Private Const Initial               As String = "Initial"
Private Const RCrN                  As String = "Copyright Notice Agree"
Private Const LastStation           As String = "Last Station"
Private Const RunInTray             As String = "Run In Tray"
Private Const Balloons              As String = "Popup Balloons"
Private Const Yeah                  As String = "Yes"
Private Const Nope                  As String = "No"
Private Const Wdrcm                 As String = "When digitally recording copyrighted material "

Private Sub AppendToWriteBuffer(ByVal Data As String, ByVal FlushBuffer As Boolean)

  Dim Lng   As Long

    Lng = Len(Data) And (hFile <> 0)
    If Lng Then
        WriteBuffer = WriteBuffer & Data
        If Len(WriteBuffer) >= MaxWriteBufferSize Or FlushBuffer Then
            If hFile Then
                lbWriting.Visible = True
                tmrOff.Enabled = True
                Print #hFile, WriteBuffer;
                NumSessionBlocksWritten = NumSessionBlocksWritten + 1
                NumBlocksWritten = NumBlocksWritten + 1
                NumBytesWritten = NumBytesWritten + Len(WriteBuffer)
                ckKill.Enabled = optSeparateFiles
                mnuKill.Visible = optSeparateFiles
                WriteBuffer = vbNullString
                If lbBuffering.Visible Then
                    If LOF(hFile) > BitrateInBits Then
                        wmPlayer.URL = OutputFilename
                        lbBuffering.Visible = False
                        btRestart.Enabled = True
                        mnuRestart.Visible = True
                        If WindowState = vbMinimized Then
                            Systray.HideBalloon
                        End If
                    End If
                End If
            End If
        End If
        NumBytesAppended = NumBytesAppended + Lng
        NumBytesAppendedForTitle = NumBytesAppendedForTitle + Lng
    End If
    Display

End Sub

Private Sub btAbout_Click()

    Load frmAbout
    With frmAbout
        .Theme = Timer Mod 27 + 1
        .AppIcon(&HFFE0C0) = Icon
        .Title(vbRed) = App.ProductName
        .Version(&HFFC0C0) = "Version " & App.Major & "." & App.Minor & "." & App.Revision
        .Copyright(vbYellow) = App.LegalCopyright
        .Otherstuff1(&HB0B0B0) = "Original Author: Final Stand Productions"
        .Otherstuff2(&H9090FF) = Wdrcm & "special regulations may apply in your country"
        .Show vbModal, Me
    End With 'FRMABOUT

    If Agreed = False Then
        Agreed = (MsgBoxEx(Wdrcm & "you must" & vbCrLf & "observe the rights of the respective copyright owners." & EndHeader & _
                 "Bitte beachten Sie die Rechte des jeweiligen Copyright-" & vbCrLf & "Eigentümers am digital aufgezeichneten Material.", _
                 vbExclamation Or vbOKCancel, "Copyright Notice / Hinweis", -1, -2, -2, 40, -56, Icon:=picIcon, OCapt:=OK & "|" & Abbrechen, NCapt:="&Agree|&Disagree") = vbOK)
        SaveSetting App.EXEName, Initial, RCrN, IIf(Agreed, Yeah, Nope)
        btRun.Enabled = Agreed
        btAbout.FontBold = Not Agreed
    End If

End Sub

Private Sub btEdit_Click()

    WaitFor Shell("notepad.exe " & fnNetStations, vbNormalFocus)

End Sub

Private Sub btRestart_Click()

    wmPlayer.URL = vbNullString
    wmPlayer.URL = OutputFilename

End Sub

Private Sub btRun_Click()

    If Running Then
        'stop reception
        Running = False
        Disconnect
        lbBuffering.Visible = False
        CloseFile
        ckUseTitleList.Enabled = True
        wmPlayer.URL = vbNullString
        lbDiscard.Visible = False
        lbDiscard.ForeColor = lbWillBe.ForeColor
        SetUI Enabled
        tmrAnim.Enabled = False
        Caption = App.ProductName
        lbStream = WaitingStation
        SetLink "", lbStream
        lbGenre = "[Unknown]"
        lbBitrate = "[Unknown]"
        lbSong.Enabled = False 'so that it does not trigger the change event
        lbSong = WaitingTitle
        SetLink "", lbSong
        lbSong.Enabled = True
        btEdit.Enabled = True
        btRestart.Enabled = True
        mnuRestart.Visible = True
        btRun.Caption = "Start &Receiving"
        mnuStartStop.Caption = "Start Receiving"
        btRun.BackColor = &HA0FFA0
        txtSave = App.Path & MusicDir & "Music" & MP3
      Else 'RUNNING = FALSE/0
        'start reception
        btRun.Enabled = False
        lbOut.Visible = False
        ckUseTitleList.Enabled = False
        Bar = 0
        If OpenFile(txtSave, optAllInOneFile) Then 'could open file
            lbInterrupted.Visible = False
            If IsConnected Then 'could open the stream
                Running = True
                btRestart.Enabled = False
                mnuRestart.Visible = False
                btEdit.Enabled = False
                On Error Resume Next
                    MkDir App.Path & MusicDir
                On Error GoTo 0
                SetUI Not Enabled
                btRun.Caption = "Stop &Receiving"
                mnuStartStop.Caption = "Stop Receiving"
                btRun.BackColor = &HA0A0FF
                ICYHeaderReceived = False
                Metablock = vbNullString
                InputBuffer = vbNullString
                AnimCounter = 0
                NumPacketsReceived = 0
                NumBytesReceived = 0
                MaxBufferSize = 0
                MaxPacketSize = 0
                NumSessionBlocksWritten = 0
                NumBytesAppended = 0
            End If
            If Not Running Then
                lbInterrupted.Visible = True
                Running = True
                btRun_Click
            End If
            tmrAnim.Enabled = True
        End If
        btRun.Enabled = True
    End If

End Sub

Private Sub btSave_Click()

    With CD
        .DialogTitle = "Enter/Select file to record..."
        .Filter = "Sound File (*" & MP3 & ")|*" & MP3
        .Flags = cdlOFNPathMustExist Or cdlOFNLongNames
        .Filename = txtSave
        .ShowSave
        On Error Resume Next
            txtSave = .Filename
            OutputFilename = .Filename
        On Error GoTo 0
    End With 'CD

End Sub

Private Sub cboStations_Click()

    CurrStation = Replace$(Stations(cboStations.ListIndex), " ", vbNullString)

End Sub

Private Function CheckTitle(Title As String) As Boolean

  'this is here tnx to michael doering

  Dim i As Long

    If UBound(Titles) And ckUseTitleList = vbChecked Then
        For i = 1 To UBound(Titles)
            If LCase(Title) Like Titles(i) Then
                CheckTitle = True
                Exit For 'loop varying i
            End If
        Next i
      Else 'NOT UBOUND(TITLES)...
        CheckTitle = True
    End If

End Function

Private Sub ckKill_Click()

    lbWillBe.Visible = (ckKill = vbChecked)
    mnuKill.Checked = (ckKill = vbChecked)
    lbDiscard.Visible = (ckKill = vbChecked Or lbDiscard.ForeColor = vbRed)

End Sub

Private Sub ckUseTitleList_Click()

    If UBound(Titles) = 0 Then
        ckUseTitleList = vbUnchecked
    End If

End Sub

Private Sub ckUseTitleList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        WaitFor Shell("notepad.exe " & fnTitles, vbNormalFocus)
    End If

End Sub

Private Sub CloseFile()

    If hFile Then
        Close hFile
        If ckKill = vbChecked Or ckStick = vbChecked Then
            On Error Resume Next
                Kill OutputFilename
                If Err = 0 Then
                    lbWillBe.Visible = False
                    lbDiscard.ForeColor = vbRed
                    lbDiscard.Visible = True
                    NumBytesAppended = NumBytesAppended - NumBytesWritten
                    NumSessionBlocksWritten = NumSessionBlocksWritten - NumBlocksWritten
                    NumBytesAppendedForTitle = 0
                    Display
                End If
            On Error GoTo 0
            ckKill = vbUnchecked
        End If
        ckKill.Enabled = False
        mnuKill.Visible = False
        hFile = 0
        btRun.Width = 111
    End If
    lbWriting.Visible = False
    NumBytesAppendedForTitle = 0
    NumBytesWritten = 0
    NumBlocksWritten = 0
    btRestart.Enabled = False
    mnuRestart.Visible = False

End Sub

Private Sub Disconnect()

    Winsock.Close
    AppendToWriteBuffer InputBuffer, True
    InputBuffer = vbNullString

End Sub

Public Sub Display()

  Dim Duration  As Long
  Dim Progress  As Long

    If NumBytesAppended < 996148 Then
        lbWritten = Format$(NumBytesAppended / 1024, "#0") & " kB"
      Else 'NOT NUMBYTESAPPENDED...
        lbWritten = Format$(NumBytesAppended / 1048576, "#0.00") & " MB"
    End If
    If BitrateInKBits Then
        Duration = NumBytesAppendedForTitle \ ByterateInBytes
        lbCurrTime(0) = Format$(Duration \ 60) & ":" & Format$(Duration Mod 60, "00")
        Duration = NumBytesAppended \ ByterateInBytes
        lbTime = Format$(Duration \ 60) & ":" & Format$(Duration Mod 60, "00")
    End If
    If optBreak(0) Then
        Progress = (Duration * 100) / Abs(Val(txtBreak(0)) * 60 + 1)
      Else 'OPTBREAK(0) = FALSE/0
        Progress = ((NumBytesAppended / 1024) * 100) / Abs(Val(txtBreak(1)) * 1024 + 1)
    End If
    If Progress > 100 Then
        Progress = 100
    End If
    Bar = Progress
    If Progress = 100 Then
        If Running Then
            btRun_Click
            lbOut.Visible = True
            Bar = 100
        End If
    End If
    Caption = App.ProductName & " [" & IIf(optBreak(0), lbTime, lbWritten) & "]"

End Sub

Private Sub FillStations()

  Dim ls        As String 'last station
  Dim i         As Long
  Dim s         As String
  Dim b         As Boolean
  Dim Cmds()    As String

    ls = GetSetting(App.EXEName, Initial, LastStation, vbNullString)
    hFile = FreeFile
    On Error Resume Next
        Open App.Path & "\" & fnNetStations For Input As hFile
        With cboStations
            If Err = 0 Then
                i = 0
                Do Until EOF(hFile)
                    Line Input #hFile, s
                    s = Trim$(s)
                    If Len(s) > 4 Then
                        If Left$(s, 1) <> ";" Then
                            If Left$(s, 1) = "*" Then
                                b = (Len(ls) = 0)
                                s = Mid$(s, 2)
                              Else 'NOT LEFT$(S,...
                                b = False
                            End If
                            Cmds = Split(s, "|")
                            If UBound(Cmds) = 0 Then
                                ReDim Preserve Cmds(0 To 1)
                                Cmds(1) = Cmds(0)
                            End If
                            .AddItem Trim$(Cmds(1))
                            ReDim Preserve Stations(0 To .NewIndex)
                            Stations(.NewIndex) = Trim$(Cmds(0))
                            If b Or .List(.NewIndex) = ls Then
                                i = .NewIndex
                            End If
                        End If
                    End If
                Loop
                Close hFile
                .ListIndex = i
              Else 'NOT ERR...
                .Text = "http://64.236.34.196:80/stream/1013" 'no station list default
                .Enabled = False
            End If
        End With 'CBOSTATIONS
        ReDim Titles(0 To 0) 'tnx michael
        Open App.Path & "\" & fnTitles For Input As hFile
        If Err = 0 Then
            i = 0
            Do Until EOF(hFile)
                Line Input #hFile, s
                s = Trim$(s)
                If Len(s) Then
                    If Left$(s, 1) <> ";" Then
                        i = i + 1
                        ReDim Preserve Titles(0 To i)
                        Titles(i) = LCase$(s)
                    End If
                End If
            Loop
            Close hFile
          Else 'NOT ERR...
            ckUseTitleList.Enabled = False
        End If
        If ckUseTitleList.Enabled = False Then
            ckUseTitleList = vbUnchecked
        End If
        hFile = 0
    On Error GoTo 0

End Sub

Private Sub Form_Activate()

    If Len(cboStations) <> 0 And (Len(txtSave) Or optSeparateFiles) And ByCommandline Then
        btRun_Click
    End If

End Sub

Private Sub Form_Initialize()

    InitCommonControls

End Sub

Private Sub Form_Load()

  Dim i         As Long
  Dim Cmds()    As String

    WaitingStation = lbStream
    WaitingTitle = lbSong

    If Not InIDE Then 'prevent the ugly focus rects
        Set EmergencyUnhook = New clsUnhook 'will unhook the subclassed hWnds on class instance destruction
        Hook WM_SETFOCUS, Array(optAllInOneFile.hWnd, _
                                optSeparateFiles.hWnd, _
                                optBreak(0).hWnd, _
                                optBreak(1).hWnd, _
                                ckKill.hWnd, _
                                ckStick.hWnd, _
                                ckUseTitleList.hWnd, _
                                btRun.hWnd, _
                                btSave.hWnd, _
                                btRestart.hWnd, _
                                btAbout.hWnd, _
                                btEdit.hWnd)
    End If
    If App.PrevInstance Then
        MsgBox App.ProductName & " is already running.", vbCritical, "Oops..."
        Unload Me
      Else 'APP.PREVINSTANCE = FALSE/0
        With App
            lbVersion = " Version " & .Major & "." & .Minor & "." & .Revision
        End With 'APP

        Set Systray = New clsSystray
        Systray.SetOwner Me

        Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
        Agreed = (GetSetting(App.EXEName, Initial, RCrN, Nope) = Yeah)
        If Agreed And GetSetting(App.EXEName, Initial, RunInTray, Nope) = Yeah Then
            WindowState = vbMinimized
        End If
        btRun.Enabled = Agreed
        btAbout.FontBold = Not Agreed
        If GetSetting(App.EXEName, Initial, Balloons, Yeah) = Yeah Then
            mnuEnaBallons_Click
          Else 'NOT GETSETTING(APP.EXENAME,...
            mnuSuppBalloons_Click
        End If
        RestartTimer AnimSpeed

        txtSave = App.Path & MusicDir & "Music" & MP3
        Caption = App.ProductName
        If Len(Trim$(Command$)) Then
            Cmds = Split(Command$, " ")
            '-a url [-s savefile] [-t timelimit] [-m sizelimit] [-o f/s]
            'eg - QuickRip.exe "-a http://64.236.34.196:80/stream/1013 -t 60 -o s"
            For i = LBound(Cmds) To UBound(Cmds) - 1 Step 2
                Select Case Cmds(i)
                  Case "-a"
                    cboStations = LCase$(Cmds(i + 1))
                    cboStations.Enabled = False
                  Case "-s"
                    txtSave = Cmds(i + 1)
                  Case "-t"
                    txtBreak(0) = Val(Cmds(i + 1))
                    optBreak(0) = True
                  Case "-m"
                    txtBreak(1) = Val(Cmds(i + 1))
                    optBreak(1) = True
                  Case "-o"
                    optSeparateFiles = (LCase$(Cmds(i + 1)) <> "f")
                    optAllInOneFile = (LCase$(Cmds(i + 1)) = "f")
                End Select
            Next i
            ByCommandline = True
          Else 'LEN(TRIM$(COMMAND$)) = FALSE/0
            FillStations
        End If
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If App.PrevInstance = False Then
        If hFile Then
            If MsgBox("You have a recording going..." & EndHeader & "Are you sure you want to quit?", vbQuestion Or vbDefaultButton2 Or vbYesNo, "About to quit") = vbYes Then
                btRun_Click
              Else 'NOT MSGBOX("YOU HAVE A RECORDING GOING..."...
                Cancel = True
            End If
        End If
        If Cancel = False Then
            On Error Resume Next
                RmDir App.Path & MusicDir 'kill dir if empty
            On Error GoTo 0
            SaveSetting App.EXEName, Initial, LastStation, cboStations.Text
            SaveSetting App.EXEName, Initial, Balloons, IIf(BalloonsEnabled, Yeah, Nope)
            If Timer Mod 32 = 0 Then 'keep nagging at odd times (once a month or so if they use it daily)
                SaveSetting App.EXEName, Initial, RCrN, Nope
                SaveSetting App.EXEName, Initial, RunInTray, Nope
              Else 'NOT TIMER...
                SaveSetting App.EXEName, Initial, RunInTray, IIf(WindowState = vbMinimized, Yeah, Nope)
            End If
        End If
    End If

End Sub

Private Sub Form_Resize()

    If WindowState = vbMinimized Then
        Hide
        DoEvents
        Systray.AddIconToTray Icon, App.ProductName & " / " & lbSong, True
        If BalloonsEnabled Then
            Systray.ShowBalloon "I am here now... (just in case you wonder)", App.ProductName, InfoIcon
        End If
        RestartTimer 2500 'to hide balloon
      Else 'NOT WINDOWSTATE...
        RestartTimer AnimSpeed
    End If

End Sub

Private Function InIDE(Optional c As Boolean = False) As Boolean

  Static b As Boolean

    b = c
    If b = False Then
        Debug.Assert InIDE(True)
    End If
    InIDE = b

End Function

Private Function IsConnected() As Boolean

  Dim RawAddress    As String
  Dim i             As Long
  Dim j             As Long
  Dim Host          As String
  Dim Port          As String
  Dim Path          As String
  Dim TimeNow       As Long

    RawAddress = Replace$(CurrStation, "http://", vbNullString, , , vbTextCompare)
    i = InStr(RawAddress, ":")
    If RawAddress <> CurrStation Or i = 0 Then
        j = InStr(i, RawAddress, "/")
        If j = 0 Then
            RawAddress = RawAddress & "/"
            j = Len(RawAddress)
        End If
        Host = Mid$(RawAddress, 1, i - 1)
        Port = CInt(Mid$(RawAddress, i + 1, j - i - 1))
        Path = "/"
        If j <> Len(RawAddress) Then
            Path = Mid$(RawAddress, j)
        End If
        Winsock.Close 'reset socket
        On Error Resume Next
            Winsock.Connect Host, Port 'try to connect
        On Error GoTo 0
        TimeNow = GetTickCount + 10000 'timeout 10 seconds
        IsConnected = True
        Screen.MousePointer = vbHourglass
        Do 'wait for connection
            DoEvents
            If Winsock.State = sckError Or GetTickCount > TimeNow Then
                Screen.MousePointer = vbNormal
                MsgBox "Failed to connect to Radio Station.", vbCritical, Caption
                IsConnected = False
            End If
        Loop Until Winsock.State = sckConnected Or IsConnected = False
        Screen.MousePointer = vbNormal

        If IsConnected Then 'send the stream request
            Winsock.SendData "GET " & Path & " HTTP/1.0" & vbCrLf & _
                             "Host: " & Host & vbCrLf & _
                             "User-Agent: WinampMPEG/2.7" & vbCrLf & _
                             "Accept: */*" & vbCrLf & _
                             "Icy-MetaData: 1" & vbCrLf & _
                             "Connection: Close" & EndHeader
        End If
      Else 'NOT RAWADDRESS...
        MsgBox UCase$(CurrStation) & EndHeader & " is an illegal Internet Resource Locator.", vbExclamation, Caption
    End If

End Function

Private Sub lbCurrTime_Change(Index As Integer)

    If Index = 0 Then 'echo
        lbCurrTime(1) = lbCurrTime(0)
    End If

End Sub

Private Sub lbSong_Change()

    If WindowState = vbMinimized Then
        If BalloonsEnabled Then
            Systray.ShowBalloon lbSong, "Currenty receiving", InfoIcon Or SoundOff
        End If
    End If
    lbSong.Tooltiptext = lbSong
    Systray.Tooltip = App.ProductName & " / " & lbSong
    If optSeparateFiles And lbSong.Enabled Then
        AppendToWriteBuffer vbNullString, True
        CloseFile
        OutputFilename = App.Path & MusicDir & lbSong & MP3
        If CheckTitle(lbSong.Caption) Then
            If OpenFile(OutputFilename, True) = False Then
                hFile = 0
            End If
          Else 'CHECKTITLE(LBSONG.CAPTION) = FALSE/0
            hFile = 0
        End If
    End If

End Sub

Private Sub lbSong_Click()

    If lbSong.Tooltiptext <> vbNullString Then
        ShellExecute 0&, vbNullString, lbSong.Tooltiptext, vbNullString, vbNullString, vbNormalFocus
    End If

End Sub

Private Sub lbStream_Change()

    lbStream.Tooltiptext = lbStream

End Sub

Private Sub lbStream_Click()

    If lbStream.Tooltiptext <> vbNullString Then
        ShellExecute 0&, vbNullString, lbStream.Tooltiptext, vbNullString, vbNullString, vbNormalFocus
    End If

End Sub

Private Sub mnuDefault_Click()

    Systray_DoubleClick vbLeftButton

End Sub

Private Sub mnuEnaBallons_Click()

    BalloonsEnabled = True
    mnuSuppBalloons.Checked = False
    mnuEnaBallons.Checked = True

End Sub

Private Sub mnuExit_Click()

    Unload Me

End Sub

Private Sub mnuKill_Click()

    With mnuKill
        .Checked = Not .Checked
        ckKill = IIf(.Checked, vbChecked, vbUnchecked)
    End With 'mnuKill

End Sub

Private Sub mnuRestart_Click()

    btRestart.Value = True

End Sub

Private Sub mnuSendMail_Click()

    With App
        SendMeMail hWnd, .ProductName & " Version " & .Major & "." & .Minor & "." & .Revision
    End With 'APP

End Sub

Private Sub mnuStartStop_Click()

    btRun.Value = btRun.Enabled

End Sub

Private Sub mnuSuppBalloons_Click()

    BalloonsEnabled = False
    mnuSuppBalloons.Checked = True
    mnuEnaBallons.Checked = False

End Sub

Private Function OpenFile(Filename As String, DoIt As Boolean) As Boolean

  'handles file open

  'opens all-in-one file unconditionally in commandline mode
  'warns the user that the all-in-one file exists in manual mode
  'appends a serial number to filename if the file exists in separate file mode

  Dim i As Long
  Dim j As Long

    OpenFile = True
    OutputFilename = Filename 'tnx coder ghost ;-)
    If DoIt Then 'this is a bypass so as not to open the all in one-file
        Do While FileExists(Filename, 0, 0)
            If optAllInOneFile Then
                If ByCommandline Then
                    Exit Do 'loop 
                  ElseIf MsgBox("File" & EndHeader & Filename & EndHeader & "already exists. Overwrite?", vbExclamation Or vbYesNo, Caption) = vbYes Then 'BYCOMMANDLINE = FALSE/0
                    Exit Do 'loop 
                  Else 'NOT MSGBOX("FILE"...
                    OpenFile = False
                    Exit Do 'loop 
                End If
              Else 'OPTALLINONEFILE = FALSE/0
                If LCase$(Right$(Filename, 5)) = "}" & MP3 Then
                    i = Len(Filename) - 5
                    j = 0
                    Do
                        Select Case True
                          Case IsNumeric(Mid$(Filename, i, 1))
                            j = -1
                          Case Mid$(Filename, i, 1) = "{"
                            Exit Do 'loop 
                          Case Else 'not numeric and not open braces
                            i = 1
                        End Select
                        i = i - 1
                    Loop While i
                End If
                If i And j Then
                    j = Val(Mid$(Filename, i + 1))
                    Filename = Replace$(Filename, "{" & j & "}" & MP3, "{" & (j + 1) & "}" & MP3, , , vbTextCompare)
                  Else 'I = FALSE/0 'NOT I...
                    Filename = Replace$(Filename, MP3, "{2}" & MP3, , , vbTextCompare)
                End If
            End If
        Loop
        If OpenFile Then
            DoEvents
            hFile = FreeFile
            On Error Resume Next
                Open Filename For Output As hFile
                OpenFile = (Err = 0)
            On Error GoTo 0
            If OpenFile Then
                btRun.Width = 84
                lbBuffering.Visible = hFile
            End If
        End If
    End If

End Function

Private Sub optAllInOneFile_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If optAllInOneFile Then
        With txtSave
            .Enabled = True
            .SetFocus
            .ForeColor = vbBlack
            .SelStart = 0
            .SelLength = .MaxLength
        End With 'TXTSAVE
        btSave.Enabled = True
    End If

End Sub

Private Sub optBreak_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    With txtBreak(Index)
        .SetFocus
        .SelStart = 0
        .SelLength = .MaxLength
    End With 'TXTBREAK(INDEX)

End Sub

Private Sub optSeparateFiles_Click()

    If optSeparateFiles Then
        txtSave.Enabled = False
        btSave.Enabled = False
    End If

End Sub

Private Sub ProcessHeader(Header As String)

  Dim Items()   As String
  Dim Parts()   As String
  Dim i         As Long
  Dim j         As Long

    If InStr(1, Header, vbCrLf, vbBinaryCompare) Then 'ICY -Header
        Items = Split(Header, vbCrLf)
        For i = 0 To UBound(Items)
            If InStr(Items(i), ":") Then
                Parts = Split(Items(i), ":")
                For j = 2 To UBound(Parts)
                    Parts(Contents) = Parts(Contents) & ":" & Parts(j)
                Next j
                Select Case Parts(Label)
                  Case "icy-url"
                    SetLink Trim$(Parts(Contents)), lbStream
                  Case "icy-br"
                    BitrateInKBits = Val(Parts(Contents))
                    BitrateInBits = BitrateInKBits * 1024
                    ByterateInBytes = BitrateInBits / 8
                    lbBitrate = BitrateInKBits & " kBit/sec"
                  Case "icy-name"
                    lbStream = Trim$(Parts(Contents))
                    lbStream.Tooltiptext = lbStream
                  Case "icy-genre"
                    lbGenre = Trim$(Parts(Contents))
                  Case "icy-metaint"
                    MetaPacketSize = Val(Parts(Contents))
                End Select
            End If
        Next i
      Else 'NOT INSTR(1,...
        Items = Split(Header, ";")
        For i = 0 To UBound(Items)
            If InStr(Items(i), "=") Then
                Parts = Split(Items(i), "=")
                Select Case LCase$(Parts(Label))
                  Case "streamtitle"
                    Parts(Contents) = Trim$(Replace$(Parts(Contents), "'", vbNullString))
                    If Len(Parts(Contents)) = 0 Then
                        lbSong = "Received " & Format$(Now, "yyyy mm dd hh nn ss")
                      Else 'NOT LEN(PARTS(CONTENTS))...
                        For j = 1 To Len(Parts(Contents)) 'remove illegal-for-filename chars
                            Select Case LCase$(Mid$(Parts(Contents), j, 1))
                              Case "0" To "9", "a" To "z", "ä", "ö", "ü", "ß" 'replace with your permitted local chars if necessary
                                'do nothing
                              Case Is < Chr$(32), ":", "\", "|", "?", "*", Is > Chr$(127)
                                Mid$(Parts(Contents), j, 1) = "#"
                              Case "&"
                                Mid$(Parts(Contents), j, 1) = "+"
                            End Select
                        Next j
                        lbSong = Parts(Contents)
                    End If
                  Case "streamurl"
                    SetLink Trim$(Replace$(Parts(Contents), "'", vbNullString)), lbSong
                End Select
            End If
        Next i
    End If

End Sub

Private Sub RestartTimer(ByVal Delay As Long)

    tmrAnim.Enabled = False
    tmrAnim.Interval = Delay
    tmrAnim.Enabled = True

End Sub

Public Sub SendMeMail(FromhWnd As Long, Subject As String)

  Dim UserName  As String
  Dim Lng       As Long

    Lng = 128
    UserName = String$(Lng, 0)
    GetUserName UserName, Lng
    UserName = Left$(UserName, Lng + (Asc(Mid$(UserName, Lng, 1)) = 0))
    If ShellExecute(FromhWnd, vbNullString, "mailto:UMGEDV@Yahoo.com?subject=" & Subject & " &body=Hi Ulli,   [your message]   Best regards from " & UserName, vbNullString, App.Path, SW_SHOWNORMAL) < SE_NO_ERROR Then
        MsgBox "Cannot send Mail from this System.", vbExclamation, "Mail disabled/not installed"
    End If

End Sub

Private Sub SetLink(Link As String, Lbl As Label)

    With Lbl
        .Tooltiptext = Link
        If Link = vbNullString Then
            .FontUnderline = False
            .MousePointer = vbDefault
          Else 'NOT LINK...
            .FontUnderline = True
            .MousePointer = 99
        End If
    End With 'LBL

End Sub

Private Sub SetUI(ByVal State As Boolean)

    txtSave.Enabled = optAllInOneFile And State
    btSave.Enabled = optAllInOneFile And State
    cboStations.Enabled = State And cboStations.ListCount
    optSeparateFiles.Enabled = State
    optAllInOneFile.Enabled = State

End Sub

Private Sub Systray_DoubleClick(Button As Integer)

    WindowState = vbNormal
    Systray.RemoveIconFromTray
    Show

End Sub

Private Sub tmrAnim_Timer()

    AnimPosn = (AnimPosn + 1) Mod 3
    picAnim(AnimPosn).ZOrder 0
    If AnimCounter = AnimTimeout Then 'timeout
        tmrAnim.Enabled = False
        lbInterrupted.Visible = Running
        btRun.Value = Running
      Else 'NOT ANIMCOUNTER...
        AnimCounter = AnimCounter + 1
    End If
    If WindowState = vbMinimized Then
        tmrAnim.Enabled = False
        Systray.HideBalloon
      ElseIf Agreed = False Then 'NOT WINDOWSTATE...
        btAbout.FontBold = Not btAbout.FontBold
    End If

End Sub

Private Sub tmrOff_Timer()

    lbWriting.Visible = False
    If lbWillBe.ForeColor <> lbDiscard.ForeColor Then
        lbDiscard.Visible = False
        lbDiscard.ForeColor = lbWillBe.ForeColor
    End If
    tmrOff.Enabled = False

End Sub

Private Sub txtSave_Change()

    txtSave.Tooltiptext = txtSave

End Sub

Private Sub txtSave_Validate(Cancel As Boolean)

    If LCase(Right$(txtSave, 4)) <> MP3 Then
        txtSave = txtSave & MP3
    End If

End Sub

Private Sub WaitFor(Task As Long)

  Dim hProcess  As Long

    Do
        hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, Task)
        If hProcess Then 'notepad is still active
            CloseHandle hProcess 'thats all we wanted to know so close the handle
        End If
        DoEvents
    Loop Until hProcess = 0
    cboStations.Clear
    FillStations

End Sub

Private Sub Winsock_DataArrival(ByVal BytesTotal As Long)

  Dim i As Long

    Winsock.GetData Packet

    NumPacketsReceived = NumPacketsReceived + 1
    lbPackets = NumPacketsReceived
    If AnimCounter = AnimTimeout Then
        tmrAnim.Enabled = True
    End If
    AnimCounter = 0 'reset timeout
    InputBuffer = InputBuffer & Packet
    NumBytesReceived = NumBytesReceived + BytesTotal
    AvgPacketSize = NumBytesReceived / NumPacketsReceived
    If Len(InputBuffer) > MaxBufferSize Then
        MaxBufferSize = Len(InputBuffer)
    End If
    If BytesTotal > MaxPacketSize Then
        MaxPacketSize = BytesTotal
    End If
    If ICYHeaderReceived Then
        If Len(InputBuffer) > MetaPacketSize Then
            i = Asc(Mid$(InputBuffer, MetaPacketSize + 1, 1)) * 16
            If i = 0 Then
                AppendToWriteBuffer Mid$(InputBuffer, 1, MetaPacketSize), False
                InputBuffer = Mid$(InputBuffer, MetaPacketSize + 2)
              Else 'NOT i...
                If MetaPacketSize + i <= Len(InputBuffer) Then
                    Metablock = Mid$(InputBuffer, MetaPacketSize + 2, i)
                    AppendToWriteBuffer Mid$(InputBuffer, 1, MetaPacketSize), True
                    InputBuffer = Mid$(InputBuffer, MetaPacketSize + i + 2)
                    ProcessHeader Metablock
                End If
            End If
        End If
      Else 'ICYHEADERRECEIVED = FALSE/0
        i = InStr(1, InputBuffer, EndHeader, vbBinaryCompare)
        If i Then
            Metablock = Mid$(InputBuffer, 1, Len(EndHeader) + i - 1)
            InputBuffer = Mid$(InputBuffer, Len(EndHeader) + i)
            ProcessHeader Metablock
            ICYHeaderReceived = True
        End If
    End If

End Sub

Private Sub wmPlayer_PlayStateChange(ByVal NewState As Long)

  'when the player overruns the buffer (which should never happen) we attempt an automatic restart

    If NewState = 1 And Not lbInterrupted.Visible Then 'restart player
        btRestart.Value = btRestart.Enabled And Running
    End If

End Sub

':) Ulli's VB Code Formatter V2.21.6 (2006-Aug-31 15:44)  Decl: 100  Code: 1021  Total: 1121 Lines
':) CommentOnly: 16 (1,4%)  Commented: 71 (6,3%)  Empty: 189 (16,9%)  Max Logic Depth: 8

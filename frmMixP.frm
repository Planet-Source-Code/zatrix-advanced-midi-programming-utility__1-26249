VERSION 5.00
Begin VB.Form frmMixP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mix Part"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   Icon            =   "frmMixP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTB 
      Height          =   2745
      Index           =   4
      Left            =   8070
      TabIndex        =   179
      Top             =   3600
      Width           =   7845
      Begin VB.CheckBox chkPorta 
         Alignment       =   1  'Right Justify
         Caption         =   "Portamento"
         Height          =   315
         Left            =   1770
         Style           =   1  'Graphical
         TabIndex        =   182
         Top             =   900
         Width           =   1125
      End
      Begin VB.PictureBox picP4 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   1800
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   180
         Top             =   1440
         Width           =   1950
      End
      Begin VB.Label lbl 
         Caption         =   "This parameter isn't included in SC-55 dumps."
         Height          =   375
         Index           =   55
         Left            =   210
         TabIndex        =   183
         Top             =   240
         Width           =   6075
      End
      Begin VB.Label lblN 
         Alignment       =   1  'Right Justify
         Caption         =   "Portamento Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   13
         Left            =   210
         TabIndex        =   181
         Tag             =   "0"
         Top             =   1440
         Width           =   1545
      End
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   60
      TabIndex        =   177
      Top             =   90
      Width           =   525
   End
   Begin VB.CommandButton cmdResetGS 
      Caption         =   "GS"
      Height          =   315
      Left            =   150
      TabIndex        =   176
      ToolTipText     =   "Reset to GS"
      Top             =   1530
      Width           =   345
   End
   Begin VB.CommandButton cmdAllNotesOff 
      Height          =   315
      Left            =   150
      Picture         =   "frmMixP.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   175
      ToolTipText     =   "All Notes Off"
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.ComboBox cmbAssign 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmMixP.frx":040C
      Left            =   5760
      List            =   "frmMixP.frx":0419
      Style           =   2  'Dropdown List
      TabIndex        =   173
      Top             =   750
      Width           =   1245
   End
   Begin VB.Frame fraPDrum 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   525
      Left            =   1620
      TabIndex        =   169
      Top             =   1230
      Visible         =   0   'False
      Width           =   2775
      Begin VB.ComboBox cmbDrumP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   170
         TabStop         =   0   'False
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lbl 
         Caption         =   "DRUM PATCH"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   54
         Left            =   30
         TabIndex        =   171
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame fraPNorm 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   525
      Left            =   870
      TabIndex        =   165
      Top             =   570
      Width           =   3555
      Begin VB.ComboBox cmbBank 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmMixP.frx":0437
         Left            =   0
         List            =   "frmMixP.frx":045C
         Style           =   2  'Dropdown List
         TabIndex        =   172
         Top             =   180
         Width           =   675
      End
      Begin VB.ComboBox cmbPatch 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   750
         Style           =   2  'Dropdown List
         TabIndex        =   166
         TabStop         =   0   'False
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lbl 
         Caption         =   "BANK"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   38
         Left            =   30
         TabIndex        =   168
         Top             =   0
         Width           =   465
      End
      Begin VB.Label lbl 
         Caption         =   "INSTRUMENT"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   39
         Left            =   780
         TabIndex        =   167
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.CheckBox chkSend 
      Height          =   315
      Left            =   150
      Picture         =   "frmMixP.frx":0484
      Style           =   1  'Graphical
      TabIndex        =   164
      ToolTipText     =   "Midi Out"
      Top             =   870
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.PictureBox picP0 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      DrawMode        =   7  'Invert
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   7
      Left            =   5940
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   158
      Top             =   2250
      Width           =   1950
   End
   Begin VB.PictureBox picP0 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      DrawMode        =   7  'Invert
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   4
      Left            =   5940
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   154
      Top             =   1260
      Width           =   1950
   End
   Begin VB.PictureBox picP0 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      DrawMode        =   7  'Invert
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   5
      Left            =   5940
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   153
      Top             =   1590
      Width           =   1950
   End
   Begin VB.PictureBox picP0 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      DrawMode        =   7  'Invert
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   6
      Left            =   5940
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   152
      Top             =   1920
      Width           =   1950
   End
   Begin VB.PictureBox picKeyR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   60
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   523
      TabIndex        =   151
      Top             =   2880
      Width           =   7840
   End
   Begin VB.CheckBox chkMono 
      Caption         =   "MONO"
      Height          =   195
      Left            =   7110
      TabIndex        =   146
      Top             =   810
      Width           =   855
   End
   Begin VB.PictureBox picP0 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      DrawMode        =   7  'Invert
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   3
      Left            =   1680
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   144
      Top             =   2250
      Width           =   1950
   End
   Begin VB.PictureBox picP0 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      DrawMode        =   7  'Invert
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   0
      Left            =   1680
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   140
      Top             =   1260
      Width           =   1950
   End
   Begin VB.PictureBox picP0 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      DrawMode        =   7  'Invert
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   1680
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   139
      Top             =   1590
      Width           =   1950
   End
   Begin VB.PictureBox picP0 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      DrawMode        =   7  'Invert
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   2
      Left            =   1680
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   138
      Top             =   1920
      Width           =   1950
   End
   Begin VB.ComboBox cmbPartMode 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmMixP.frx":0586
      Left            =   4410
      List            =   "frmMixP.frx":0593
      Style           =   2  'Dropdown List
      TabIndex        =   136
      Top             =   750
      Width           =   1125
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   60
      TabIndex        =   121
      Top             =   3300
      Width           =   7875
      Begin VB.OptionButton optTab 
         Caption         =   "Extra"
         Height          =   315
         Index           =   4
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   178
         Top             =   30
         Width           =   1470
      End
      Begin VB.OptionButton optTab 
         Caption         =   "Controllers"
         Height          =   315
         Index           =   3
         Left            =   4770
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   30
         Width           =   1500
      End
      Begin VB.OptionButton optTab 
         Caption         =   "Switches"
         Height          =   315
         Index           =   2
         Left            =   3180
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   30
         Width           =   1500
      End
      Begin VB.OptionButton optTab 
         Caption         =   "Tuning"
         Height          =   315
         Index           =   1
         Left            =   1590
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   30
         Width           =   1500
      End
      Begin VB.OptionButton optTab 
         Caption         =   "Vib / Env"
         Height          =   315
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   30
         Value           =   -1  'True
         Width           =   1500
      End
   End
   Begin VB.Frame fraTB 
      Height          =   2745
      Index           =   3
      Left            =   8160
      TabIndex        =   67
      Top             =   2610
      Width           =   7845
      Begin VB.OptionButton optCTL 
         Caption         =   "CC2"
         Height          =   285
         Index           =   5
         Left            =   5730
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   210
         Width           =   855
      End
      Begin VB.OptionButton optCTL 
         Caption         =   "CC1"
         Height          =   285
         Index           =   4
         Left            =   4860
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   210
         Width           =   855
      End
      Begin VB.OptionButton optCTL 
         Caption         =   "PAf"
         Height          =   285
         Index           =   3
         Left            =   3990
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   210
         Width           =   855
      End
      Begin VB.OptionButton optCTL 
         Caption         =   "CAf"
         Height          =   285
         Index           =   2
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   210
         Width           =   855
      End
      Begin VB.OptionButton optCTL 
         Caption         =   "Bend"
         Height          =   285
         Index           =   1
         Left            =   2250
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   210
         Width           =   855
      End
      Begin VB.OptionButton optCTL 
         Caption         =   "Mod"
         Height          =   285
         Index           =   0
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   210
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.PictureBox picP3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   10
         Left            =   5370
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   84
         Top             =   2310
         Width           =   1950
      End
      Begin VB.PictureBox picP3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   9
         Left            =   5370
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   83
         Top             =   1980
         Width           =   1950
      End
      Begin VB.PictureBox picP3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   8
         Left            =   5370
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   82
         Top             =   1650
         Width           =   1950
      End
      Begin VB.PictureBox picP3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   7
         Left            =   5370
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   81
         Top             =   1320
         Width           =   1950
      End
      Begin VB.PictureBox picP3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   6
         Left            =   1620
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   79
         Top             =   2310
         Width           =   1950
      End
      Begin VB.PictureBox picP3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   5
         Left            =   1620
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   77
         Top             =   1980
         Width           =   1950
      End
      Begin VB.PictureBox picP3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   4
         Left            =   1620
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   75
         Top             =   1650
         Width           =   1950
      End
      Begin VB.PictureBox picP3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   3
         Left            =   1620
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   73
         Top             =   1320
         Width           =   1950
      End
      Begin VB.PictureBox picP3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   5370
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   72
         Top             =   600
         Width           =   1950
      End
      Begin VB.PictureBox picP3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   1620
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   70
         Top             =   930
         Width           =   1950
      End
      Begin VB.PictureBox picP3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   1620
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   68
         Top             =   600
         Width           =   1950
      End
      Begin VB.Label lbl 
         Caption         =   "AMPLITUDE  CTL"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   22
         Left            =   4110
         TabIndex        =   135
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "RATE CTL"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   27
         Left            =   4350
         TabIndex        =   134
         Top             =   1380
         Width           =   930
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "PITCH DEP"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   28
         Left            =   4350
         TabIndex        =   133
         Top             =   1710
         Width           =   930
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "TVF DEPTH"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   29
         Left            =   4350
         TabIndex        =   132
         Top             =   2040
         Width           =   930
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "TVA DEPTH"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   30
         Left            =   4350
         TabIndex        =   131
         Top             =   2370
         Width           =   930
      End
      Begin VB.Label lbl 
         Caption         =   "LFO2"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   32
         Left            =   3990
         TabIndex        =   86
         Top             =   1860
         Width           =   555
      End
      Begin VB.Label lbl 
         Caption         =   "LFO1"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   31
         Left            =   270
         TabIndex        =   85
         Top             =   1860
         Width           =   555
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "TVA DEPTH"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   26
         Left            =   600
         TabIndex        =   80
         Top             =   2370
         Width           =   945
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "TVF DEPTH"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   25
         Left            =   600
         TabIndex        =   78
         Top             =   2040
         Width           =   945
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "PITCH DEP"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   24
         Left            =   600
         TabIndex        =   76
         Top             =   1710
         Width           =   945
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "RATE CTL"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   23
         Left            =   600
         TabIndex        =   74
         Top             =   1380
         Width           =   945
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "TVF CUTOFF CTL"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   21
         Left            =   600
         TabIndex        =   71
         Top             =   990
         Width           =   945
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "PITCH  CTL"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   20
         Left            =   600
         TabIndex        =   69
         Top             =   660
         Width           =   945
      End
   End
   Begin VB.Frame fraTB 
      Height          =   2745
      Index           =   2
      Left            =   8160
      TabIndex        =   17
      Top             =   1800
      Width           =   7845
      Begin VB.OptionButton optChan 
         Caption         =   "Off"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   16
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   2100
         Width           =   480
      End
      Begin VB.OptionButton optChan 
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   15
         Left            =   6030
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   2100
         Width           =   330
      End
      Begin VB.OptionButton optChan 
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   14
         Left            =   5700
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   2100
         Width           =   330
      End
      Begin VB.OptionButton optChan 
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         Left            =   5370
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   2100
         Width           =   330
      End
      Begin VB.OptionButton optChan 
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   2100
         Width           =   330
      End
      Begin VB.OptionButton optChan 
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   4710
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   2100
         Width           =   330
      End
      Begin VB.OptionButton optChan 
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   2100
         Width           =   330
      End
      Begin VB.OptionButton optChan 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   4050
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   2100
         Width           =   330
      End
      Begin VB.OptionButton optChan 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   2100
         Width           =   330
      End
      Begin VB.OptionButton optChan 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   3390
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   2100
         Width           =   330
      End
      Begin VB.OptionButton optChan 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   3060
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   2100
         Width           =   330
      End
      Begin VB.OptionButton optChan 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   2730
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   2100
         Width           =   330
      End
      Begin VB.OptionButton optChan 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   2100
         Width           =   330
      End
      Begin VB.OptionButton optChan 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   2070
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   2100
         Width           =   330
      End
      Begin VB.OptionButton optChan 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1740
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   2100
         Width           =   330
      End
      Begin VB.OptionButton optChan 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1410
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   2100
         Width           =   330
      End
      Begin VB.OptionButton optChan 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   2100
         Value           =   -1  'True
         Width           =   330
      End
      Begin VB.CheckBox chkRx 
         Caption         =   "Soft"
         Height          =   195
         Index           =   15
         Left            =   5670
         TabIndex        =   33
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkRx 
         Caption         =   "Sostenuto"
         Height          =   195
         Index           =   14
         Left            =   5670
         TabIndex        =   32
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkRx 
         Caption         =   "Portamento"
         Height          =   195
         Index           =   13
         Left            =   5670
         TabIndex        =   31
         Top             =   840
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkRx 
         Caption         =   "Hold1"
         Height          =   195
         Index           =   12
         Left            =   5670
         TabIndex        =   30
         Top             =   600
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkRx 
         Caption         =   "Expression"
         Height          =   195
         Index           =   11
         Left            =   4170
         TabIndex        =   29
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkRx 
         Caption         =   "Panpot"
         Height          =   195
         Index           =   10
         Left            =   4170
         TabIndex        =   28
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkRx 
         Caption         =   "Volume"
         Height          =   195
         Index           =   9
         Left            =   4170
         TabIndex        =   27
         Top             =   840
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkRx 
         Caption         =   "Modulation"
         Height          =   195
         Index           =   8
         Left            =   4170
         TabIndex        =   26
         Top             =   600
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkRx 
         Caption         =   "NRPN"
         Height          =   195
         Index           =   7
         Left            =   2610
         TabIndex        =   25
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkRx 
         Caption         =   "RPN"
         Height          =   195
         Index           =   6
         Left            =   2610
         TabIndex        =   24
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkRx 
         Caption         =   "Note Messages"
         Height          =   195
         Index           =   5
         Left            =   2610
         TabIndex        =   23
         Top             =   840
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkRx 
         Caption         =   "Poly Pressure"
         Height          =   195
         Index           =   4
         Left            =   2610
         TabIndex        =   22
         Top             =   600
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkRx 
         Caption         =   "Contr. Change"
         Height          =   195
         Index           =   3
         Left            =   1080
         TabIndex        =   21
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkRx 
         Caption         =   "Progr. Change"
         Height          =   195
         Index           =   2
         Left            =   1080
         TabIndex        =   20
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkRx 
         Caption         =   "Ch. Pressure"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   19
         Top             =   840
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkRx 
         Caption         =   "Pitch Bend"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   18
         Top             =   600
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.Label lbl 
         Caption         =   "RX CHANNEL"
         Height          =   225
         Index           =   37
         Left            =   1080
         TabIndex        =   130
         Top             =   1830
         Width           =   1455
      End
      Begin VB.Label lbl 
         Caption         =   "RX SWITCHES"
         Height          =   225
         Index           =   36
         Left            =   1080
         TabIndex        =   129
         Top             =   270
         Width           =   1455
      End
   End
   Begin VB.Frame fraTB 
      Height          =   2745
      Index           =   1
      Left            =   8130
      TabIndex        =   42
      Top             =   750
      Width           =   7845
      Begin VB.PictureBox picOF 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5190
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   149
         Top             =   1770
         Width           =   1950
      End
      Begin VB.PictureBox picKeySh 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5190
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   147
         Top             =   1110
         Width           =   1950
      End
      Begin VB.PictureBox picP2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         ForeColor       =   &H00FF0000&
         Height          =   1950
         Index           =   11
         Left            =   4080
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   65
         Top             =   630
         Width           =   255
      End
      Begin VB.PictureBox picP2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         ForeColor       =   &H00FF0000&
         Height          =   1950
         Index           =   10
         Left            =   3780
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   63
         Top             =   630
         Width           =   255
      End
      Begin VB.PictureBox picP2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         ForeColor       =   &H00FF0000&
         Height          =   1950
         Index           =   9
         Left            =   3480
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   61
         Top             =   630
         Width           =   255
      End
      Begin VB.PictureBox picP2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         ForeColor       =   &H00FF0000&
         Height          =   1950
         Index           =   8
         Left            =   3180
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   59
         Top             =   630
         Width           =   255
      End
      Begin VB.PictureBox picP2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         ForeColor       =   &H00FF0000&
         Height          =   1950
         Index           =   7
         Left            =   2880
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   57
         Top             =   630
         Width           =   255
      End
      Begin VB.PictureBox picP2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         ForeColor       =   &H00FF0000&
         Height          =   1950
         Index           =   6
         Left            =   2580
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   55
         Top             =   630
         Width           =   255
      End
      Begin VB.PictureBox picP2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         ForeColor       =   &H00FF0000&
         Height          =   1950
         Index           =   5
         Left            =   2280
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   53
         Top             =   630
         Width           =   255
      End
      Begin VB.PictureBox picP2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         ForeColor       =   &H00FF0000&
         Height          =   1950
         Index           =   4
         Left            =   1980
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   51
         Top             =   630
         Width           =   255
      End
      Begin VB.PictureBox picP2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         ForeColor       =   &H00FF0000&
         Height          =   1950
         Index           =   3
         Left            =   1680
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   49
         Top             =   630
         Width           =   255
      End
      Begin VB.PictureBox picP2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         ForeColor       =   &H00FF0000&
         Height          =   1950
         Index           =   2
         Left            =   1380
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   47
         Top             =   630
         Width           =   255
      End
      Begin VB.PictureBox picP2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         ForeColor       =   &H00FF0000&
         Height          =   1950
         Index           =   1
         Left            =   1080
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   45
         Top             =   630
         Width           =   255
      End
      Begin VB.PictureBox picP2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         ForeColor       =   &H00FF0000&
         Height          =   1950
         Index           =   0
         Left            =   780
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   43
         Top             =   630
         Width           =   255
      End
      Begin VB.Label lbl 
         Caption         =   "PITCH OFFS FINE"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   46
         Left            =   5190
         TabIndex        =   150
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lbl 
         Caption         =   "PITCH KEY SHIFT"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   45
         Left            =   5190
         TabIndex        =   148
         Top             =   900
         Width           =   1185
      End
      Begin VB.Label lbl 
         Caption         =   "SCALE TUNING"
         Height          =   225
         Index           =   35
         Left            =   1950
         TabIndex        =   128
         Top             =   180
         Width           =   1305
      End
      Begin VB.Label lbl 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   19
         Left            =   4140
         TabIndex        =   66
         Top             =   420
         Width           =   195
      End
      Begin VB.Label lbl 
         Caption         =   "A#"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   18
         Left            =   3810
         TabIndex        =   64
         Top             =   420
         Width           =   195
      End
      Begin VB.Label lbl 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   17
         Left            =   3540
         TabIndex        =   62
         Top             =   420
         Width           =   195
      End
      Begin VB.Label lbl 
         Caption         =   "G#"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   16
         Left            =   3210
         TabIndex        =   60
         Top             =   420
         Width           =   195
      End
      Begin VB.Label lbl 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   15
         Left            =   2940
         TabIndex        =   58
         Top             =   420
         Width           =   195
      End
      Begin VB.Label lbl 
         Caption         =   "F#"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   14
         Left            =   2610
         TabIndex        =   56
         Top             =   420
         Width           =   195
      End
      Begin VB.Label lbl 
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   13
         Left            =   2370
         TabIndex        =   54
         Top             =   420
         Width           =   195
      End
      Begin VB.Label lbl 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   12
         Left            =   2070
         TabIndex        =   52
         Top             =   420
         Width           =   195
      End
      Begin VB.Label lbl 
         Caption         =   "D#"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   11
         Left            =   1710
         TabIndex        =   50
         Top             =   420
         Width           =   195
      End
      Begin VB.Label lbl 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   10
         Left            =   1440
         TabIndex        =   48
         Top             =   420
         Width           =   195
      End
      Begin VB.Label lbl 
         Caption         =   "C#"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   9
         Left            =   1110
         TabIndex        =   46
         Top             =   420
         Width           =   195
      End
      Begin VB.Label lbl 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   8
         Left            =   840
         TabIndex        =   44
         Top             =   420
         Width           =   135
      End
   End
   Begin VB.Frame fraTB 
      Height          =   2745
      Index           =   0
      Left            =   60
      TabIndex        =   34
      Top             =   3630
      Width           =   7845
      Begin VB.PictureBox picP1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   3
         Left            =   2400
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   95
         Top             =   2340
         Width           =   1545
      End
      Begin VB.PictureBox picP1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   4
         Left            =   5160
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   94
         Top             =   2340
         Width           =   1545
      End
      Begin VB.PictureBox picP1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   5
         Left            =   2400
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   90
         Top             =   1290
         Width           =   1545
      End
      Begin VB.PictureBox picP1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   6
         Left            =   2400
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   89
         Top             =   1620
         Width           =   1545
      End
      Begin VB.PictureBox picP1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   7
         Left            =   2400
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   88
         Top             =   1950
         Width           =   1545
      End
      Begin VB.PictureBox picEnv 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawWidth       =   2
         ForeColor       =   &H80000008&
         Height          =   945
         Left            =   4080
         ScaleHeight     =   61
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   200
         TabIndex        =   87
         Top             =   1290
         Width           =   3030
      End
      Begin VB.PictureBox picVib 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   945
         Left            =   4080
         ScaleHeight     =   61
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   200
         TabIndex        =   41
         Top             =   240
         Width           =   3030
      End
      Begin VB.PictureBox picP1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   2400
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   39
         Top             =   900
         Width           =   1545
      End
      Begin VB.PictureBox picP1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   2400
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   37
         Top             =   570
         Width           =   1545
      End
      Begin VB.PictureBox picP1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         DrawMode        =   7  'Invert
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   2400
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   35
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         Caption         =   " ENVELOPE"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   34
         Left            =   600
         TabIndex        =   127
         Top             =   1650
         Width           =   975
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         Caption         =   " VIBRATO"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   33
         Left            =   600
         TabIndex        =   126
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "CUTOFF"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   7
         Left            =   1680
         TabIndex        =   97
         Top             =   2400
         Width           =   645
      End
      Begin VB.Label lbl 
         Caption         =   "RESON."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   6
         Left            =   4590
         TabIndex        =   96
         Top             =   2400
         Width           =   585
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "ATTACK"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   5
         Left            =   1680
         TabIndex        =   93
         Top             =   1350
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "DECAY"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   4
         Left            =   1680
         TabIndex        =   92
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "RELEASE"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   3
         Left            =   1680
         TabIndex        =   91
         Top             =   2010
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "DELAY"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   2
         Left            =   1680
         TabIndex        =   40
         Top             =   960
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "DEPTH"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   1
         Left            =   1680
         TabIndex        =   38
         Top             =   630
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "RATE"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   1680
         TabIndex        =   36
         Top             =   300
         Width           =   645
      End
   End
   Begin VB.Frame fraPart 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   660
      TabIndex        =   0
      Top             =   90
      Width           =   7305
      Begin VB.OptionButton optPart 
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   15
         Left            =   6780
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Width           =   450
      End
      Begin VB.OptionButton optPart 
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   14
         Left            =   6330
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         Width           =   450
      End
      Begin VB.OptionButton optPart 
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         Width           =   450
      End
      Begin VB.OptionButton optPart 
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         Left            =   5430
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         Width           =   450
      End
      Begin VB.OptionButton optPart 
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   4980
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   450
      End
      Begin VB.OptionButton optPart 
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   4530
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   450
      End
      Begin VB.OptionButton optPart 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   450
      End
      Begin VB.OptionButton optPart 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   3630
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   450
      End
      Begin VB.OptionButton optPart 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   3180
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   450
      End
      Begin VB.OptionButton optPart 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   2730
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   450
      End
      Begin VB.OptionButton optPart 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   450
      End
      Begin VB.OptionButton optPart 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   1830
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   450
      End
      Begin VB.OptionButton optPart 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   450
      End
      Begin VB.OptionButton optPart 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   930
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   450
      End
      Begin VB.OptionButton optPart 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   450
      End
      Begin VB.OptionButton optPart 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   450
      End
   End
   Begin VB.Label lbl 
      Caption         =   "ASSIGN MODE"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   53
      Left            =   5790
      TabIndex        =   174
      Top             =   570
      Width           =   1095
   End
   Begin VB.Label lblKeyRH 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "H - C#9"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   7080
      TabIndex        =   163
      ToolTipText     =   "Use Control + Right mouse button"
      Top             =   2610
      Width           =   675
   End
   Begin VB.Label lblKeyRL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L - C#-1"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   360
      TabIndex        =   162
      ToolTipText     =   "Use Control + Left mouse button"
      Top             =   2610
      Width           =   675
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "CONTROLLER NUMBER"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   52
      Left            =   4200
      TabIndex        =   161
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "VELOCITY SENSE"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   51
      Left            =   4380
      TabIndex        =   160
      Top             =   1410
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "CC2"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   50
      Left            =   5490
      TabIndex        =   159
      Top             =   2310
      Width           =   345
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "DEPTH"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   49
      Left            =   5250
      TabIndex        =   157
      Top             =   1320
      Width           =   585
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "OFFSET"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   48
      Left            =   5220
      TabIndex        =   156
      Top             =   1650
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "CC1"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   47
      Left            =   5490
      TabIndex        =   155
      Top             =   1980
      Width           =   345
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "CHORUS"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   44
      Left            =   870
      TabIndex        =   145
      Top             =   2310
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "VOLUME"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   43
      Left            =   870
      TabIndex        =   143
      Top             =   1320
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "PAN"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   42
      Left            =   870
      TabIndex        =   142
      Top             =   1650
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "REVERB"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   41
      Left            =   870
      TabIndex        =   141
      Top             =   1980
      Width           =   705
   End
   Begin VB.Label lbl 
      Caption         =   "PART MODE"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   40
      Left            =   4440
      TabIndex        =   137
      Top             =   570
      Width           =   975
   End
   Begin VB.Shape shpKeyRH 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   2790
      Width           =   75
   End
   Begin VB.Shape shpKeyRL 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   540
      Shape           =   3  'Circle
      Top             =   2790
      Width           =   75
   End
End
Attribute VB_Name = "frmMixP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form is independant of the data in sc55par.dat
'           is easy to extract from the project

Option Explicit
Dim DoSend As Boolean   ' true = user changed value, so send
                        ' false = program changed value
Dim Send As Boolean     ' send to sound canvas

Dim CurTab As Long      ' current group tab
Dim CurPart As Integer  ' current part
Dim CurCTL As Long      ' group 3 current contr. option
Dim CurKey As Byte      ' remember note on --> note off
Dim dmpPart As Long     ' current part dump order
Dim baseOffs As Long    ' offset into part dump bytes
Dim Syx1 As String      ' start F0 41 10 42 12 40 of SysEx commandstring
                                    
Dim expar(1, 16) As Integer ' extra param, not found in dump

' key range & play
Public Sub MakePiano(pic As PictureBox)
   Dim wX1 As Long, wY1 As Long
   Dim wdX As Long, wdY As Long
   Dim zX1 As Long, zY1 As Long
   Dim zdX As Long, zdY As Long
   Dim AaWTs As Long                   ' count white keys
   Dim I As Long                       ' counter
   
   wX1 = 0: wY1 = 0: wdX = 7: wdY = 22 ' witte toets
   zX1 = 5: zY1 = 0: zdX = 4: zdY = 16 ' zwarte

   AaWTs = (128 / 12) * 7

   pic.Width = AaWTs * wdX * 15
   pic.AutoRedraw = True
   
   ' make 1st white key & copy other white keys
   pic.Line (wX1, wY1)-Step(wdX, wdY), QBColor(15), BF
   pic.Line (wX1, wY1)-Step(wdX, wdY), QBColor(0), B
   For I = 0 To AaWTs - 1
      BitBlt pic.hDC, wX1 + I * wdX, wY1, wdX, wdY + 1, pic.hDC, wX1, wY1, SRCCOPY
   Next I
      
   ' 1st black & copy other
   pic.Line (zX1, zY1)-Step(zdX, zdY), QBColor(0), BF
   For I = 1 To AaWTs - 1
      If Mid("110111", (I Mod 7) + 1, 1) = "1" Then
         BitBlt pic.hDC, zX1 + I * wdX, zY1, zdX + 1, zdY, pic.hDC, zX1, zY1, SRCCOPY
         End If
   Next I
   
   pic.Line (pic.ScaleWidth - 1, wY1)-Step(0, wdY), QBColor(0)
   pic.Picture = pic.Image
   pic.AutoRedraw = False
End Sub

' group3 - controllers
Private Sub SetCurCTL(ByVal Index As Integer, ByVal Value As Single)
   Dim BO As Long          ' bytes offset - first
   Dim jA As Byte          ' junoir address byte
   Dim rI As Long          ' real index BO
   Dim rValue As Single    ' real value
   Dim pCaption As String
   Dim X1 As Long, X2 As Long
   Dim midX As Long        ' positioning caption
   Dim ComStr As String    ' SysEx commandstring

   Select Case Index
   Case 0: ' pitch
      X1 = 0: rValue = Value - 64
      X2 = rValue
      If CurCTL = 1 Then midX = 12 Else midX = 0
   Case 1
      rValue = Convert(Value, &H0, &H7F, -9600, 9600) ' cutoff
      X1 = 64
   Case 2
      rValue = Convert(Value, &H0, &H7F, -100, 100)  ' ampli
      X1 = 64
   Case 3, 7
      rValue = Convert(Value, &H0, &H7F, -10, 10) ' lfo rate
      X1 = 64
   Case 4, 8
      rValue = Convert(Value, &H0, &H7F, 0, 600) ' lfo pitch depth
      X1 = 0
   Case 5, 9
      rValue = Convert(Value, &H0, &H7F, 0, 2400) ' lfo tvf depth
      X1 = 0
   Case 6, 10
      rValue = Convert(Value, &H0, &H7F, 0, 100) ' lfo tva depth
      X1 = 0
   End Select
   If Index > 0 Then X2 = Value: midX = 64
   pCaption = IIf(CInt(rValue) = rValue, Format(rValue), Format(rValue, "#0.0"))
   picP3(Index).Cls
   picP3(Index).CurrentX = midX - picP3(Index).TextWidth(pCaption) / 2
   picP3(Index).CurrentY = -2
   picP3(Index).Print pCaption
   picP3(Index).Line (X1, 0)-(X2, 17), QBColor(14), BF

   If DoSend = False Then Exit Sub
   ' mod, bnd, Caf, Paf, cc1, cc2
   BO = Choose(CurCTL + 1, 40, 52, 64, 76, 88, 100)
   If Index > 2 Then rI = Index + 1 Else rI = Index ' one row skipped each
   dmpB(baseOffs + BO + rI) = CByte(Value)
   
   If Send = False Then Exit Sub
   jA = CurCTL * &H10 + Index
   ComStr = Syx1 & Chr(&H20 Or dmpPart) & Chr(jA) & Chr(Value)
   ComStr = ComStr & Chr(-(&H40 + &H20 + dmpPart + jA + Value) And 127)
   ComStr = ComStr & Chr(&HF7)
   SysExDT1 ComStr

End Sub

' group 0 - volume, panpot, ...
Private Sub SetG0(ByVal Index As Integer, ByVal X As Single)
   Dim BO As Long
   Dim ComStr As String
   Dim pCaption As String
   
   If Index = 1 Then pCaption = isPanPot(X) Else pCaption = CStr(X)
   picP0(Index).Cls
   picP0(Index).CurrentX = 64 - picP0(Index).TextWidth(pCaption) / 2
   picP0(Index).CurrentY = -2
   picP0(Index).Print pCaption
   picP0(Index).Line (0, 0)-(X, 17), QBColor(14), BF
   
   If DoSend = False Then Exit Sub
   BO = Choose(Index + 1, 8, 9, 15, 14, 10, 11, 38, 39)
   dmpB(baseOffs + BO) = X
   
   If Send = False Then Exit Sub
   Select Case Index
      Case 0 ' volume
         midiMessageOut = CONTROLLER_CHANGE + CurPart
         midiData1 = MAIN_VOLUME
         midiData2 = CByte(X)
         SendMidiShortOut
      Case 1 ' pan
         midiMessageOut = CONTROLLER_CHANGE + CurPart
         midiData1 = PAN
         midiData2 = CByte(X)
         SendMidiShortOut
      Case 2 ' reverb
         midiMessageOut = CONTROLLER_CHANGE + CurPart
         midiData1 = &H5B
         midiData2 = CByte(X)
         SendMidiShortOut
      Case 3 ' chorus
         midiMessageOut = CONTROLLER_CHANGE + CurPart
         midiData1 = &H5D
         midiData2 = CByte(X)
         SendMidiShortOut
      Case 4 ' velo dep
         ComStr = Syx1 & Chr(&H10 Or dmpPart) & Chr(&H1A) & Chr(X)
         ComStr = ComStr & Chr(-(&H40 + &H10 + dmpPart + &H1A + X) And 127)
         ComStr = ComStr & Chr(&HF7)
         SysExDT1 ComStr
      Case 5 ' velo offs
         ComStr = Syx1 & Chr(&H10 Or dmpPart) & Chr(&H1B) & Chr(X)
         ComStr = ComStr & Chr(-(&H40 + &H10 + dmpPart + &H1B + X) And 127)
         ComStr = ComStr & Chr(&HF7)
         SysExDT1 ComStr
      Case 6 ' CC1
         ComStr = Syx1 & Chr(&H10 Or dmpPart) & Chr(&H1F) & Chr(X)
         ComStr = ComStr & Chr(-(&H40 + &H10 + dmpPart + &H1F + X) And 127)
         ComStr = ComStr & Chr(&HF7)
         SysExDT1 ComStr
      Case 7 ' CC2
         ComStr = Syx1 & Chr(&H10 Or dmpPart) & Chr(&H20) & Chr(X)
         ComStr = ComStr & Chr(-(&H40 + &H10 + dmpPart + &H20 + X) And 127)
         ComStr = ComStr & Chr(&HF7)
         SysExDT1 ComStr
   End Select

End Sub

' group 1 - vibrato/envelope cutoff/resonance
Private Sub SetG1(ByVal Index As Integer, ByVal X As Single)
   Dim BO As Long
   Dim Value As Byte
   
   picP1(Index).Cls
   picP1(Index).CurrentX = 50 - picP1(Index).TextWidth(CStr(X - 50)) / 2
   picP1(Index).CurrentY = -2
   picP1(Index).Print CStr(X - 50)
   picP1(Index).Line (50, 0)-(X, 17), QBColor(14), BF
   
   If DoSend = False Then Exit Sub
   ' vib rate, vib dep, vib del, cutoff, resonance, attack, decay, release
   BO = Choose(Index + 1, 16, 17, 23, 18, 19, 20, 21, 22)
   Value = &HE + X
   dmpB(baseOffs + BO) = Value
   
   If Send = False Then Exit Sub
   Select Case Index
      Case 0 ' vib rate
         midiMessageOut = CONTROLLER_CHANGE + CurPart
         midiData1 = &H63: midiData2 = &H1: SendMidiShortOut
         midiData1 = &H62: midiData2 = &H8: SendMidiShortOut
         midiData1 = &H6: midiData2 = CByte(Value): SendMidiShortOut
      Case 1 ' vib dep
         midiMessageOut = CONTROLLER_CHANGE + CurPart
         midiData1 = &H63: midiData2 = &H1: SendMidiShortOut
         midiData1 = &H62: midiData2 = &H9: SendMidiShortOut
         midiData1 = &H6: midiData2 = CByte(Value): SendMidiShortOut
      Case 2 ' vib delay
         midiMessageOut = CONTROLLER_CHANGE + CurPart
         midiData1 = &H63: midiData2 = &H1: SendMidiShortOut
         midiData1 = &H62: midiData2 = &HA: SendMidiShortOut
         midiData1 = &H6: midiData2 = CByte(Value): SendMidiShortOut
      Case 3 ' cutoff
         midiMessageOut = CONTROLLER_CHANGE + CurPart
         midiData1 = &H63: midiData2 = &H1: SendMidiShortOut
         midiData1 = &H62: midiData2 = &H20: SendMidiShortOut
         midiData1 = &H6: midiData2 = CByte(Value): SendMidiShortOut
      Case 4 ' reson
         midiMessageOut = CONTROLLER_CHANGE + CurPart
         midiData1 = &H63: midiData2 = &H1: SendMidiShortOut
         midiData1 = &H62: midiData2 = &H21: SendMidiShortOut
         midiData1 = &H6: midiData2 = CByte(Value): SendMidiShortOut
      Case 5 ' attack
         midiMessageOut = CONTROLLER_CHANGE + CurPart
         midiData1 = &H63: midiData2 = &H1: SendMidiShortOut
         midiData1 = &H62: midiData2 = &H63: SendMidiShortOut
         midiData1 = &H6: midiData2 = CByte(Value): SendMidiShortOut
      Case 6 ' decay
         midiMessageOut = CONTROLLER_CHANGE + CurPart
         midiData1 = &H63: midiData2 = &H1: SendMidiShortOut
         midiData1 = &H62: midiData2 = &H64: SendMidiShortOut
         midiData1 = &H6: midiData2 = CByte(Value): SendMidiShortOut
      Case 7 ' release
         midiMessageOut = CONTROLLER_CHANGE + CurPart
         midiData1 = &H63: midiData2 = &H1: SendMidiShortOut
         midiData1 = &H62: midiData2 = &H66: SendMidiShortOut
         midiData1 = &H6: midiData2 = CByte(Value): SendMidiShortOut
   End Select

End Sub

' group 2 - scale tuning
Private Sub SetG2(ByVal Index As Integer, ByVal Y As Single)
   Dim I As Long
   Dim ComStr As String
   Dim sum As Byte
   Dim iB As Byte
   
   picP2(Index).Cls
   picP2(Index).Line (0, 64)-(17, Y), QBColor(14), BF
   
   If DoSend = False Then Exit Sub
   dmpB(baseOffs + 26 + Index) = 127 - Y
   
   If Send = False Then Exit Sub
   ComStr = Syx1 & Chr(&H10 Or dmpPart) & Chr(&H40)
   sum = &H40 + &H10 + dmpPart + &H40
   For I = 0 To 11
      iB = dmpB(baseOffs + 26 + I)
      sum = CByte(CInt(sum + CInt(iB)) Mod 256)
      ComStr = ComStr & Chr(iB)
   Next I
   ComStr = ComStr & Chr(-(sum) And 127)
   ComStr = ComStr & Chr(&HF7)
   SysExDT1 ComStr

End Sub

' group 4 - extra param.
Private Sub SetG4(ByVal Index As Integer, ByVal X As Single)
   Dim BO As Long
   picP4(Index).Cls
   picP4(Index).CurrentX = 64 - picP4(Index).TextWidth(CStr(X)) / 2
   picP4(Index).CurrentY = -2
   picP4(Index).Print CStr(X)
   picP4(Index).Line (0, 0)-(X, 17), QBColor(14), BF
   expar(Index, CurPart) = X
   If Send = False Then Exit Sub
   Select Case Index
      Case 0: ' porta. time
         midiMessageOut = CONTROLLER_CHANGE + CurPart
         midiData1 = PORTAMENTO_TIME
         midiData2 = CByte(X)
         SendMidiShortOut
      Case 1: ' modulation
         midiMessageOut = CONTROLLER_CHANGE + CurPart
         midiData1 = MOD_WHEEL
         midiData2 = CByte(X)
         SendMidiShortOut
      Case 2: ' pitch bend
         midiMessageOut = PITCH_BEND + CurPart
         midiData1 = CByte(X And &H7F)
         midiData2 = CByte((X \ 256) And &H7F)
         SendMidiShortOut
   End Select

End Sub

Private Sub SetKeyRange(ByVal No As Long, ByVal LowHigh As String)
   Dim Oct As Long
   Dim Key As String
   Dim mX As Long, lblX As Long
   Dim BO As Byte
   Dim ComStr As String, jA As Byte
   
   Oct = No \ 12
   Key = isNote(No)
   mX = Oct * 49 + Choose((No Mod 12) + 1, 0, 0.5, 1, 1.5, 2, 3, 3.5, 4, 4.5, 5, 5.5, 6, 7) * 7
   lblX = (mX * 15) - lblKeyRL.Width / 2
   If UCase(Left(LowHigh, 1)) = "L" Then
      lblKeyRL.Caption = "Lo" & " " & Key
      lblKeyRL.Left = IIf(lblX > picKeyR.Left, lblX, picKeyR.Left)
      shpKeyRL.Left = (mX + 6) * 15
      BO = 12: jA = &H1D
      Else
      lblKeyRH.Caption = "Hi" & " " & Key
      lblKeyRH.Left = IIf(lblX + lblKeyRL.Width < picKeyR.Left + picKeyR.Width, lblX, picKeyR.Left + picKeyR.Width - lblKeyRH.Width)
      shpKeyRH.Left = (mX + 6) * 15
      BO = 13: jA = &H1E
      End If
      
   If DoSend = False Then Exit Sub
   dmpB(baseOffs + BO) = No
   
   If Send = False Then Exit Sub
   ComStr = Syx1 & Chr(&H10 Or dmpPart) & Chr(jA) & Chr(No)
   ComStr = ComStr & Chr(-(&H40 + &H10 + dmpPart + jA + No) And 127)
   ComStr = ComStr & Chr(&HF7)
   SysExDT1 ComStr

End Sub

Private Sub SetKeyShift(ByVal X As Single)
   Dim oB As Byte
   Dim Value As Long
   Dim pCaption As String
   Dim ComStr As String
   
   oB = 64 + X
   dmpB(baseOffs + 6) = oB
   Value = -24 + (oB - &H28)
   pCaption = CStr(Value)
   picKeySh.Cls
   picKeySh.CurrentX = -picKeySh.TextWidth(pCaption) / 2
   picKeySh.CurrentY = -2
   picKeySh.Print pCaption
   picKeySh.Line (0, 0)-(Value, 17), QBColor(14), BF
   
   If Send = False Then Exit Sub
   ComStr = Syx1 & Chr(&H10 Or dmpPart) & Chr(&H16) & Chr(oB)
   ComStr = ComStr & Chr(-(&H40 + &H10 + dmpPart + &H16 + oB) And 127)
   ComStr = ComStr & Chr(&HF7)
   SysExDT1 ComStr
End Sub

Private Sub SetOF(ByVal X As Single)
   Dim oB As Byte             ' out byte
   Dim jN As Byte, sN As Byte ' junior & senior nibbles
   Dim Value As Long          ' real value
   Dim pCaption As String     ' param caption
   Dim ComStr As String       ' sysex commandstring
   
   oB = 128 + X
   Value = -120 + (oB - &H8)
   pCaption = IIf(Value / 10 = Value \ 10, Format(Value \ 10), Format(Value / 10, "#0.0"))
   picOF.Cls
   picOF.CurrentX = -picOF.TextWidth(pCaption) / 2
   picOF.CurrentY = -2
   picOF.Print pCaption
   picOF.Line (0, 0)-(Value, 17), QBColor(14), BF
   
   If DoSend = False Then Exit Sub
   dmpB(baseOffs + 7) = oB
   
   If Send = False Then Exit Sub
   jN = oB And &HF
   sN = (oB \ 16) And &HF
   ComStr = Syx1 & Chr(&H10 Or dmpPart) & Chr(&H17) & Chr(sN) & Chr(jN)
   ComStr = ComStr & Chr(-(&H40 + &H10 + dmpPart + &H17 + jN + sN) And 127)
   ComStr = ComStr & Chr(&HF7)
   SysExDT1 ComStr
End Sub

' used in ShowCurPart and when current controller option changes
Private Sub ShowCurCTL()
   Dim BO As Long ' bytes offset - first
   Dim I As Long, rI As Long

   ' mod, bnd, Caf, Paf, cc1, cc2
   BO = Choose(CurCTL + 1, 40, 52, 64, 76, 88, 100)
   If CurCTL = 1 Then picP3(0).Scale (0, 0)-(24, 17) Else picP3(0).Scale (-24, 0)-(24, 17)
   For I = 0 To 10
      If I > 2 Then rI = I + 1 Else rI = I ' one row skipped each
      SetCurCTL I, dmpB(baseOffs + BO + rI)
   Next I
End Sub

Private Sub ShowCurPart()
   Dim iB As Byte, iB2 As Byte, iB3 As Byte
   Dim BO As Long ' bytes offset
   Dim I As Long, Lng As Long
   Dim Value As Single
   Dim pCaption As String
   
   DoSend = False
   
   ' bank, prg nr, mode, mono/poly, assign
   iB = dmpB(baseOffs + 5)  ' partmode BO = 5
   iB2 = dmpB(baseOffs + 0) ' bank BO = 0
   iB3 = dmpB(baseOffs + 1) ' prg nr BO = 1
   Value = (iB And &H60) \ 32
   cmbPartMode.ListIndex = Value
   If Value = 0 Then
      ' partmode normal
      For I = 0 To cmbBank.ListCount - 1
         If Val(cmbBank.List(I)) = iB2 Then cmbBank.ListIndex = I: Exit For
      Next I
      cmbPatch.ListIndex = iB3
      Else
      ' partmode drum map
      For I = 0 To cmbDrumP.ListCount - 1
         If cmbDrumP.ItemData(I) = iB3 Then cmbDrumP.ListIndex = I: Exit For
      Next I
      End If
   chkMono.Value = IIf((iB And &H80) = 0, 1, 0) ' mono/poly
   cmbAssign.ListIndex = iB And 3 ' assign mode
   
   ' group 0
   For I = 0 To 7
      ' volume, pan, reverb, chorus, velo dep, velo offs, cc1, cc2
      BO = Choose(I + 1, 8, 9, 15, 14, 10, 11, 38, 39)
      iB = dmpB(baseOffs + BO)
      SetG0 I, iB
   Next I
   
   ' group 1 - vibr / env
   For I = 0 To 7
      ' vib rate, vib dep, vib del, cutoff, resonance, attack, decay, release
      BO = Choose(I + 1, 16, 17, 23, 18, 19, 20, 21, 22)
      SetG1 I, dmpB(baseOffs + BO) - &HE
   Next I
   ShowVibr
   ShowEnvelope
   
   ' group 2 - scale tuning
   For I = 0 To 11
      SetG2 I, 127 - dmpB(baseOffs + 26 + I)
   Next I
   
   ShowCurCTL                                ' group 3 - lfo
   optChan(dmpB(baseOffs + 4)).Value = True  ' channel
   SetKeyRange dmpB(baseOffs + 12), "Low"    ' key range
   SetKeyRange dmpB(baseOffs + 13), "High"
   GetRxSwitches                             ' Rx. Switches
   SetKeyShift (dmpB(baseOffs + 6) - 64)     ' keyshift
   SetOF (dmpB(baseOffs + 7) - 128)          ' pitch offset fine
   SetG4 0, expar(0, CurPart)                ' portamento time
   chkPorta.Value = expar(1, CurPart)        ' portamento on/off
   DoSend = True
End Sub

Private Sub ShowEnvelope()
   Dim AT As Long
   Dim DC As Long
   Dim RL As Long
   Dim X(4) As Long, Y(4) As Long
   
   AT = -50 + dmpB(baseOffs + 20) - &HE
   DC = -50 + dmpB(baseOffs + 21) - &HE
   RL = -50 + dmpB(baseOffs + 22) - &HE

   If AT < 0 Then
      X(1) = 50 + AT: Y(1) = 100
   ElseIf AT > 0 Then
      X(1) = 50: Y(1) = 100 - AT
   ElseIf AT = 0 Then
      X(1) = 50: Y(1) = 100
   End If
   If DC < 0 Then
      X(2) = 100 + DC: Y(2) = Y(1) - 50
   ElseIf DC > 0 Then
      X(2) = 100: Y(2) = Y(1) - 50 + DC
   ElseIf DC = 0 Then
      X(2) = 100: Y(2) = Y(1) - 50
   End If
   X(3) = 150: Y(3) = Y(2)
   If RL < 0 Then
      X(4) = 200 + RL: Y(4) = 0
   ElseIf RL > 0 Then
      X(4) = 200: Y(4) = RL
   ElseIf RL = 0 Then
      X(4) = 200: Y(4) = 0
   End If
   
   picEnv.Cls
   picEnv.Line -(X(1), Y(1))
   picEnv.Line -(X(2), Y(2))
   picEnv.Line -(X(3), Y(3))
   picEnv.Line -(X(4), Y(4))
      
End Sub

Private Sub ShowVibr()
   Dim RT As Long
   Dim DP As Long
   Dim DL As Long
   Dim X As Long, Y As Long
   Dim X1 As Long, Y1 As Long
   Dim Ang As Long
   
   RT = (dmpB(baseOffs + 16) - 64) * 127 / 180
   DP = dmpB(baseOffs + 17) - 64
   DL = dmpB(baseOffs + 23)
   picVib.Cls
   picVib.CurrentX = 10
   picVib.CurrentY = 0
   X1 = 10: Y1 = 0
   Ang = 0
   For X = 0 To 180 Step 2
      If X > DL Then
         Y = DP * Sin(Ang)
         Ang = Ang + RT
         Else
         Y = 0
         End If
      picVib.Line -(X1 + X, Y)
   Next X
End Sub

Private Sub chkMono_Click()
   Dim iB As Byte, Value As Byte
   Dim ComStr As String
   
   If DoSend = False Then Exit Sub
   iB = dmpB(baseOffs + 5)
   iB = iB Or &H80
   If chkMono.Value = 1 Then iB = iB Xor &H80: Value = 0 Else Value = 1
   dmpB(baseOffs + 5) = iB
   
   If Send = False Then Exit Sub
   ComStr = Syx1 & Chr(&H10 Or dmpPart) & Chr(&H13) & Chr(Value)
   ComStr = ComStr & Chr(-(&H40 + &H10 + dmpPart + &H13 + Value) And 127)
   ComStr = ComStr & Chr(&HF7)
   SysExDT1 ComStr
   
End Sub

Private Sub chkPorta_Click()
   If DoSend = False Then Exit Sub
   expar(1, CurPart) = chkPorta.Value
   If Send = False Then Exit Sub
   midiMessageOut = CONTROLLER_CHANGE + CurPart
   midiData1 = PORTAMENTO
   If chkPorta.Value = 1 Then
      midiData2 = CByte(127)
      Else
      midiData2 = CByte(0)
      End If
   SendMidiShortOut
End Sub

Private Sub chkRx_Click(Index As Integer)
   SetRxSwitches Index
End Sub

Private Sub chkSend_Click()
   If chkSend.Value = 1 Then
      If hMidiOUT = 0 Then MidiOUT_Port "open"
      If hMidiOUT = 0 Then chkSend.Value = 0
      Send = True
      Else
      Send = False
      End If
End Sub

Private Sub cmbAssign_Click()
   Dim iB As Byte
   Dim ComStr As String
   
   If DoSend = False Then Exit Sub
   iB = dmpB(baseOffs + 5)
   iB = (iB And &HFC) Or cmbAssign.ListIndex
   dmpB(baseOffs + 5) = iB
   
   If Send = False Then Exit Sub
   ComStr = Syx1 & Chr(&H10 Or dmpPart) & Chr(&H14) & Chr(cmbAssign.ListIndex)
   ComStr = ComStr & Chr(-(&H40 + &H10 + dmpPart + &H14 + cmbAssign.ListIndex) And 127)
   ComStr = ComStr & Chr(&HF7)
   SysExDT1 ComStr
End Sub

Private Sub cmbBank_Click()
   Dim iB As Byte
   Dim I As Long, LI As Integer
   Dim Bnk As Integer
   
   Screen.MousePointer = vbHourglass
   Bnk = Val(cmbBank.List(cmbBank.ListIndex))
   dmpB(baseOffs + 0) = CByte(Bnk)
   If Send = True And DoSend = True Then
      midiMessageOut = &HB0 Or CurPart
      midiData1 = 0
      midiData2 = Bnk
      SendMidiShortOut
      midiMessageOut = &HB0 Or CurPart
      midiData1 = &H20
      midiData2 = 0
      SendMidiShortOut
      End If
      
   LI = cmbPatch.ListIndex
   cmbPatch.Enabled = False
   cmbPatch.Clear
   For I = 0 To 127
      cmbPatch.AddItem isPatch(Bnk, I)
   Next I
   cmbPatch.Enabled = True
   cmbPatch.ListIndex = LI ' automaticaly a program change msg
   ' will be send to the sound module as it should, ohterwise
   ' the bank change wouldn't come through.
   Screen.MousePointer = vbDefault

End Sub

Private Sub cmbDrumP_Click()
   If DoSend = False Then Exit Sub
   dmpB(baseOffs + 1) = CByte(cmbDrumP.ItemData(cmbDrumP.ListIndex))
   If Send = True Then
      midiMessageOut = PROGRAM_CHANGE + CurPart
      midiData1 = CByte(cmbDrumP.ItemData(cmbDrumP.ListIndex))
      midiData2 = 0
      SendMidiShortOut
      End If

End Sub

Private Sub cmbPartMode_Click()
   Dim iB As Byte, PM As Byte
   iB = dmpB(baseOffs + 5)
   iB = setBit(iB, 5, False)
   iB = setBit(iB, 6, False)
   PM = cmbPartMode.ListIndex
   If PM = 0 Then
      fraPDrum.Visible = False
      fraPNorm.Visible = True
      cmbPatch.ListIndex = 0
      Else
      fraPNorm.Visible = False
      fraPDrum.Visible = True
      iB = iB Or (cmbPartMode.ListIndex * 32)
      cmbDrumP.ListIndex = 0
      End If
   dmpB(baseOffs + 5) = iB
   If Not (Send = True And DoSend = True) Then Exit Sub
   Dim ComStr As String
   ComStr = Syx1 & Chr(&H10 Or dmpPart) & Chr(&H15) & Chr(PM)
   ComStr = ComStr & Chr(-(&H40 + &H10 + dmpPart + &H15 + PM) And 127)
   ComStr = ComStr & Chr(&HF7)
   SysExDT1 ComStr
   midiMessageOut = PROGRAM_CHANGE + CurPart
   midiData1 = 0
   midiData2 = 0
   SendMidiShortOut

End Sub

Private Sub GetRxSwitches()
   Dim B1 As Byte, B2 As Byte
   ' Rx. switches
   B1 = dmpB(baseOffs + 2)
   B2 = dmpB(baseOffs + 3)
   chkRx(0).Value = getBit(B1, 6)  ' 1 Rx.Bend  10111111 11111111
   chkRx(1).Value = getBit(B1, 5)  ' 2 Rx.Caf   11011111 11111111
   chkRx(2).Value = getBit(B1, 4)  ' 3 Rx.PrgCh 11101111 11111111
   chkRx(3).Value = getBit(B1, 3)  ' 4 Rx.CtlCh 11110111 11111111
   chkRx(4).Value = getBit(B1, 2)  ' 5 Rx.Paf   11111011 11111111
   chkRx(5).Value = getBit(B1, 1)  ' 6 Rx.Note  11111101 11111111
   chkRx(6).Value = getBit(B2, 0)  ' 7 Rx.RPN   11111111 11111110
   chkRx(7).Value = getBit(B1, 7)  ' 8 Rx.NRPN  01111111 11111111
   chkRx(8).Value = getBit(B2, 1)  ' 9 Rx.Modul 11111111 11111101
   chkRx(9).Value = getBit(B2, 2)  ' 0 Rx.Volum 11111111 11111011
   chkRx(10).Value = getBit(B2, 3) '10 Rx.Pan   11111111 11110111
   chkRx(11).Value = getBit(B2, 4) '11 Rx.Expr  11111111 11101111
   chkRx(12).Value = getBit(B2, 5) '12 Rx.Hold  11111111 11011111
   chkRx(13).Value = getBit(B2, 6) '13 Rx.Port  11111111 10111111
   chkRx(14).Value = getBit(B2, 7) '14 Rx.Sost  11111111 01111111
   chkRx(15).Value = getBit(B1, 0) '15 Rx.Soft  11111110 11111111
End Sub
Private Sub SetRxSwitches(ByVal Index As Integer)
   Dim B1 As Byte, B2 As Byte, jA As Byte
   Dim SetTo As Boolean
   Dim ComStr As String
   
   If DoSend = False Then Exit Sub
   SetTo = IIf(chkRx(Index).Value = 1, True, False)
   ' Rx. switches
   B1 = dmpB(baseOffs + 2)
   B2 = dmpB(baseOffs + 3)
   Select Case Index
   Case 0: B1 = setBit(B1, 6, SetTo)
   Case 1: B1 = setBit(B1, 5, SetTo)
   Case 2: B1 = setBit(B1, 4, SetTo)
   Case 3: B1 = setBit(B1, 3, SetTo)
   Case 4: B1 = setBit(B1, 2, SetTo)
   Case 5: B1 = setBit(B1, 1, SetTo)
   Case 6: B2 = setBit(B2, 0, SetTo)
   Case 7: B1 = setBit(B1, 7, SetTo)
   Case 8: B2 = setBit(B2, 1, SetTo)
   Case 9: B2 = setBit(B2, 2, SetTo)
   Case 10: B2 = setBit(B2, 3, SetTo)
   Case 11: B2 = setBit(B2, 4, SetTo)
   Case 12: B2 = setBit(B2, 5, SetTo)
   Case 13: B2 = setBit(B2, 6, SetTo)
   Case 14: B2 = setBit(B2, 7, SetTo)
   Case 15: B1 = setBit(B1, 0, SetTo)
   End Select
   dmpB(baseOffs + 2) = B1
   dmpB(baseOffs + 3) = B2

   If Send = False Then Exit Sub
   jA = &H3 + Index
   ComStr = Syx1 & Chr(&H10 Or dmpPart) & Chr(jA) & Chr(chkRx(Index).Value)
   ComStr = ComStr & Chr(-(&H40 + &H10 + dmpPart + jA + chkRx(Index).Value) And 127)
   ComStr = ComStr & Chr(&HF7)
   SysExDT1 ComStr

End Sub

Private Sub cmbPatch_Click()
   If DoSend = False Then Exit Sub
   dmpB(baseOffs + 1) = CByte(cmbPatch.ListIndex)
   If Send = True Then
      midiMessageOut = PROGRAM_CHANGE + CurPart
      midiData1 = CByte(cmbPatch.ListIndex)
      midiData2 = 0
      SendMidiShortOut
      End If
End Sub

Private Sub cmdAll_Click()
   frmMixA.Show
End Sub

Private Sub cmdAllNotesOff_Click()
   If Send = False Then Exit Sub
   midiMessageOut = CONTROLLER_CHANGE + CurPart
   midiData1 = &H7B
   midiData2 = CByte(0)
   SendMidiShortOut
End Sub

Private Sub cmdResetGS_Click()
   GSResetAll
   ShowCurPart
End Sub

Private Sub Form_Load()
   Dim I As Long
   Dim txt As String
   
   Me.Move 0, 0, 8100, 6900
   fraPDrum.Top = fraPNorm.Top
   picKeySh.Scale (-24, 0)-(24, 17)
   picOF.Scale (-120, 0)-(120, 17)
   picVib.Scale (0, 63)-(200, -64)
   picEnv.Scale (0, 100)-(200, 0)
   
   ' Tabs
   For I = 1 To 4
      fraTB(I).Move fraTB(0).Left, fraTB(0).Top
      fraTB(I).Visible = False
   Next I
   
   ' patches bank 0
   For I = 0 To 127
      cmbPatch.AddItem isPatch(0, I)
   Next I
   cmbPatch.ListIndex = 0
   
   ' drum patches
   For I = 0 To 127
      txt = isDrumSet(I)
      If txt <> "" Then
         cmbDrumP.AddItem txt
         cmbDrumP.ItemData(cmbDrumP.NewIndex) = I
         End If
   Next I
   cmbDrumP.ListIndex = 0
   cmbBank.ListIndex = 0
   cmbPartMode.ListIndex = 0
   
   Syx1 = Chr(&HF0) & Chr(&H41) & Chr(&H10) & Chr(&H42) & Chr(&H12) & Chr(&H40)
   MakePiano picKeyR
   
   optPart(CurPart).Value = True
   If hMidiOUT <> 0 Then chkSend.Value = 1
End Sub

Private Sub optChan_Click(Index As Integer)
   Dim ComStr As String
   If DoSend = False Then Exit Sub
   dmpB(baseOffs + 4) = CByte(Index)
   If Send = False Then Exit Sub
   ComStr = Syx1 & Chr(&H10 Or dmpPart) & Chr(&H2) & Chr(Index)
   ComStr = ComStr & Chr(-(&H40 + &H10 + dmpPart + &H2 + Index) And 127)
   ComStr = ComStr & Chr(&HF7)
   SysExDT1 ComStr
End Sub

Private Sub optCTL_Click(Index As Integer)
   CurCTL = Index
   ShowCurCTL
End Sub

Private Sub optPart_Click(Index As Integer)
   CurPart = Index
   SetChannel Index
   dmpPart = Choose(CurPart + 1, 1, 2, 3, 4, 5, 6, 7, 8, 9, 0, 10, 11, 12, 13, 14, 15)
   baseOffs = 72 + dmpPart * 112
   ShowCurPart
End Sub

Private Sub optTab_Click(Index As Integer)
   fraTB(CurTab).Visible = False
   CurTab = Index
   fraTB(CurTab).Visible = True
End Sub

Private Sub picKeyR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim Oct As Long
   Dim No As Long
   Dim mX As Single
   
   Oct = X \ 49
   If picKeyR.Point(X, Y) = 0 And Y < 17 Then
      mX = X - 4
      No = Oct * 12 + Choose(((mX \ 7) Mod 7) + 1, 1, 3, 5, 6, 8, 10, 11)
      Else
      mX = X
      No = Oct * 12 + Choose(((mX \ 7) Mod 7) + 1, 0, 2, 4, 5, 7, 9, 11)
      End If
   
   If Shift = 0 Then
      If Send = True Then
         midiMessageOut = NOTE_ON + CurPart
         midiData1 = No
         midiData2 = 100
         SendMidiShortOut
         CurKey = No
         End If
      Else
      SetKeyRange No, IIf(Button = 1, "Low", "High")
      End If
End Sub

Private Sub picKeyR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Send = True And Shift = 0 Then
      midiMessageOut = NOTE_OFF + CurPart
      midiData1 = CurKey
      midiData2 = 100
      SendMidiShortOut
      End If
End Sub

Private Sub picKeySh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetKeyShift X
End Sub

Private Sub picKeySh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 0 Then Exit Sub
   SetKeyShift X
End Sub

Private Sub picOF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetOF X
End Sub

Private Sub picOF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 0 Then Exit Sub
   SetOF X
End Sub

Private Sub picP0_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetG0 Index, X
End Sub

Private Sub picP0_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 0 Then Exit Sub
   If X < 0 Or X > 127 Then Exit Sub
   SetG0 Index, X
End Sub

Private Sub picP1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetG1 Index, X
End Sub

Private Sub picP1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 0 Then Exit Sub
   If X < 0 Or X > 100 Then Exit Sub
   SetG1 Index, X
   If Index >= 0 And Index <= 2 Then ShowVibr
   If Index >= 5 And Index <= 7 Then ShowEnvelope
End Sub

Private Sub picP1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Index >= 0 And Index <= 2 Then ShowVibr
   If Index >= 5 And Index <= 7 Then ShowEnvelope
End Sub

Private Sub picP2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 0 Then Exit Sub
   If Y < 0 Or Y > 127 Then Exit Sub
   picP2(Index).Cls
   picP2(Index).Line (0, 64)-(17, Y), QBColor(14), BF
End Sub

Private Sub picP2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Y < 0 Then Y = 0
   If Y > 127 Then Y = 127
   SetG2 Index, Y
End Sub

Private Sub picP3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim mX As Single
   If Index = 0 Then mX = Int(&H40 + X) Else mX = X
   SetCurCTL Index, mX
End Sub

Private Sub picP3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim mX As Single
   If Button = 0 Then Exit Sub
   If Index = 0 Then mX = Int(&H40 + X) Else mX = X
   If X < picP3(Index).ScaleLeft Or X > picP3(Index).ScaleLeft + picP3(Index).ScaleWidth Then Exit Sub
   SetCurCTL Index, mX
End Sub

Private Sub picP4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetG4 Index, X
End Sub

Private Sub picP4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 0 Then Exit Sub
   If X < 0 Or X > 127 Then Exit Sub
   SetG4 Index, X
End Sub


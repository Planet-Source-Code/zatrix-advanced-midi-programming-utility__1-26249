VERSION 5.00
Begin VB.Form frmMixA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mix All"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   Icon            =   "frmMixA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSend 
      Height          =   315
      Left            =   3780
      Picture         =   "frmMixA.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Midi Out"
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.HScrollBar HSMas 
      Height          =   255
      Index           =   3
      Left            =   5040
      Max             =   24
      Min             =   -24
      TabIndex        =   54
      Top             =   390
      Value           =   -12
      Width           =   2175
   End
   Begin VB.HScrollBar HSMas 
      Height          =   255
      Index           =   2
      Left            =   5040
      Max             =   1000
      Min             =   -1000
      TabIndex        =   52
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.HScrollBar HSCho 
      Height          =   255
      Index           =   6
      LargeChange     =   16
      Left            =   990
      Max             =   127
      TabIndex        =   47
      Top             =   3960
      Value           =   1
      Width           =   2175
   End
   Begin VB.HScrollBar HSCho 
      Height          =   255
      Index           =   5
      LargeChange     =   16
      Left            =   990
      Max             =   127
      TabIndex        =   42
      Top             =   3630
      Value           =   1
      Width           =   2175
   End
   Begin VB.HScrollBar HSCho 
      Height          =   255
      Index           =   4
      LargeChange     =   16
      Left            =   990
      Max             =   127
      TabIndex        =   40
      Top             =   3300
      Value           =   1
      Width           =   2175
   End
   Begin VB.HScrollBar HSCho 
      Height          =   255
      Index           =   3
      LargeChange     =   16
      Left            =   990
      Max             =   127
      TabIndex        =   38
      Top             =   2970
      Value           =   1
      Width           =   2175
   End
   Begin VB.HScrollBar HSRev 
      Height          =   255
      Index           =   5
      LargeChange     =   16
      Left            =   5040
      Max             =   127
      TabIndex        =   30
      Top             =   3630
      Value           =   1
      Width           =   2175
   End
   Begin VB.HScrollBar HSRev 
      Height          =   255
      Index           =   4
      LargeChange     =   16
      Left            =   5040
      Max             =   127
      TabIndex        =   28
      Top             =   3300
      Value           =   1
      Width           =   2175
   End
   Begin VB.HScrollBar HSRev 
      Height          =   255
      Index           =   3
      LargeChange     =   16
      Left            =   5040
      Max             =   127
      TabIndex        =   26
      Top             =   2970
      Value           =   1
      Width           =   2175
   End
   Begin VB.ComboBox cmbChoMac 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmMixA.frx":040C
      Left            =   990
      List            =   "frmMixA.frx":040E
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   1500
      Width           =   1875
   End
   Begin VB.HScrollBar HSCho 
      Height          =   255
      Index           =   1
      LargeChange     =   16
      Left            =   990
      Max             =   127
      TabIndex        =   19
      Top             =   2310
      Value           =   1
      Width           =   2175
   End
   Begin VB.HScrollBar HSCho 
      Height          =   255
      Index           =   0
      Left            =   990
      Max             =   7
      TabIndex        =   17
      Top             =   1980
      Value           =   1
      Width           =   2175
   End
   Begin VB.HScrollBar HSCho 
      Height          =   255
      Index           =   2
      LargeChange     =   16
      Left            =   990
      Max             =   127
      TabIndex        =   15
      Top             =   2640
      Value           =   1
      Width           =   2175
   End
   Begin VB.HScrollBar HSRev 
      Height          =   255
      Index           =   1
      Left            =   5040
      Max             =   7
      TabIndex        =   13
      Top             =   2310
      Value           =   1
      Width           =   2175
   End
   Begin VB.HScrollBar HSRev 
      Height          =   255
      Index           =   0
      Left            =   5040
      Max             =   7
      TabIndex        =   11
      Top             =   1980
      Value           =   1
      Width           =   2175
   End
   Begin VB.HScrollBar HSRev 
      Height          =   255
      Index           =   2
      LargeChange     =   16
      Left            =   5040
      Max             =   127
      TabIndex        =   9
      Top             =   2640
      Value           =   1
      Width           =   2175
   End
   Begin VB.HScrollBar HSMas 
      Height          =   255
      Index           =   1
      LargeChange     =   16
      Left            =   990
      Max             =   127
      Min             =   1
      TabIndex        =   7
      Top             =   720
      Value           =   1
      Width           =   2175
   End
   Begin VB.HScrollBar HSMas 
      Height          =   255
      Index           =   0
      LargeChange     =   16
      Left            =   990
      Max             =   127
      TabIndex        =   2
      Top             =   390
      Width           =   2175
   End
   Begin VB.ComboBox cmbRevMac 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmMixA.frx":0410
      Left            =   5040
      List            =   "frmMixA.frx":0412
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1500
      Width           =   1875
   End
   Begin VB.Line Line2 
      X1              =   3930
      X2              =   3930
      Y1              =   1140
      Y2              =   4560
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8670
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "KEYSHIFT"
      Height          =   195
      Index           =   21
      Left            =   4170
      TabIndex        =   57
      Top             =   390
      Width           =   825
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "TUNING"
      Height          =   195
      Index           =   20
      Left            =   4170
      TabIndex        =   56
      Top             =   750
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblValMas 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   7230
      TabIndex        =   55
      Top             =   360
      Width           =   705
   End
   Begin VB.Label lblValMas 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   7230
      TabIndex        =   53
      Top             =   690
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "PANPOT"
      Height          =   195
      Index           =   19
      Left            =   30
      TabIndex        =   51
      Top             =   750
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "VOLUME"
      Height          =   195
      Index           =   18
      Left            =   30
      TabIndex        =   50
      Top             =   390
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "SEND LEV TO REV"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   17
      Left            =   30
      TabIndex        =   49
      Top             =   3930
      Width           =   915
   End
   Begin VB.Label lblValCho 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   3180
      TabIndex        =   48
      Top             =   3930
      Width           =   675
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "DELAY"
      Height          =   195
      Index           =   16
      Left            =   30
      TabIndex        =   46
      Top             =   3000
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "DEPTH"
      Height          =   195
      Index           =   15
      Left            =   30
      TabIndex        =   45
      Top             =   3660
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "RATE"
      Height          =   195
      Index           =   14
      Left            =   30
      TabIndex        =   44
      Top             =   3330
      Width           =   915
   End
   Begin VB.Label lblValCho 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   3180
      TabIndex        =   43
      Top             =   3600
      Width           =   675
   End
   Begin VB.Label lblValCho 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   3180
      TabIndex        =   41
      Top             =   3270
      Width           =   675
   End
   Begin VB.Label lblValCho 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   3180
      TabIndex        =   39
      Top             =   2940
      Width           =   675
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "FEEDBACK"
      Height          =   195
      Index           =   13
      Left            =   30
      TabIndex        =   37
      Top             =   2670
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "LEVEL"
      Height          =   195
      Index           =   12
      Left            =   30
      TabIndex        =   36
      Top             =   2340
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "PRE-LPF"
      Height          =   195
      Index           =   11
      Left            =   30
      TabIndex        =   35
      Top             =   2010
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "TIME"
      Height          =   195
      Index           =   10
      Left            =   4020
      TabIndex        =   34
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "SEND LEV TO CHO"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   4020
      TabIndex        =   33
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "FEEDBACK"
      Height          =   195
      Index           =   8
      Left            =   4020
      TabIndex        =   32
      Top             =   3330
      Width           =   975
   End
   Begin VB.Label lblValRev 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   7230
      TabIndex        =   31
      Top             =   3600
      Width           =   705
   End
   Begin VB.Label lblValRev 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   7230
      TabIndex        =   29
      Top             =   3270
      Width           =   705
   End
   Begin VB.Label lblValRev 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   7230
      TabIndex        =   27
      Top             =   2940
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "LEVEL"
      Height          =   195
      Index           =   7
      Left            =   4020
      TabIndex        =   25
      Top             =   2670
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "PRE-LPF"
      Height          =   195
      Index           =   6
      Left            =   4020
      TabIndex        =   24
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "CHARACTER"
      Height          =   195
      Index           =   5
      Left            =   4020
      TabIndex        =   23
      Top             =   2010
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "MACRO"
      Height          =   195
      Index           =   4
      Left            =   30
      TabIndex        =   22
      Top             =   1590
      Width           =   915
   End
   Begin VB.Label lblValCho 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   3180
      TabIndex        =   20
      Top             =   2280
      Width           =   675
   End
   Begin VB.Label lblValCho 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   3180
      TabIndex        =   18
      Top             =   1950
      Width           =   675
   End
   Begin VB.Label lblValCho 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   3180
      TabIndex        =   16
      Top             =   2610
      Width           =   675
   End
   Begin VB.Label lblValRev 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   7230
      TabIndex        =   14
      Top             =   2280
      Width           =   705
   End
   Begin VB.Label lblValRev 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   7230
      TabIndex        =   12
      Top             =   1950
      Width           =   705
   End
   Begin VB.Label lblValRev 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   7230
      TabIndex        =   10
      Top             =   2610
      Width           =   705
   End
   Begin VB.Label lblValMas 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   3180
      TabIndex        =   8
      Top             =   690
      Width           =   675
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "MACRO"
      Height          =   195
      Index           =   3
      Left            =   4020
      TabIndex        =   6
      Top             =   1590
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "CHORUS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   5
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label lbl 
      Caption         =   "MASTER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3570
      TabIndex        =   4
      Top             =   30
      Width           =   825
   End
   Begin VB.Label lblValMas 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   3180
      TabIndex        =   3
      Top             =   360
      Width           =   675
   End
   Begin VB.Label lbl 
      Caption         =   "REVERB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   7140
      TabIndex        =   1
      Top             =   1170
      Width           =   795
   End
End
Attribute VB_Name = "frmMixA"
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

Dim Syx1 As String

Private Sub InitALL()
   HSMas(0).Value = dmpB(2)
   HSMas(1).Value = dmpB(6)
   'hsmas(2).Value = dmpB(0)
   HSMas(3).Value = IIf(dmpB(5) = 0, 0, 64 - dmpB(5)) '64-??

   cmbRevMac.ListIndex = dmpB(42)
   HSRev(0).Value = dmpB(43)
   HSRev(1).Value = dmpB(44)
   HSRev(2).Value = dmpB(45)
   HSRev(3).Value = dmpB(46)
   HSRev(4).Value = dmpB(47)
   HSRev(5).Value = dmpB(48)
   
   cmbChoMac.ListIndex = dmpB(50)
   HSCho(0).Value = dmpB(51)
   HSCho(1).Value = dmpB(52)
   HSCho(2).Value = dmpB(53)
   HSCho(3).Value = dmpB(54)
   HSCho(4).Value = dmpB(55)
   HSCho(5).Value = dmpB(56)
   HSCho(6).Value = dmpB(57)
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

Private Sub cmbChoMac_Click()
   Dim LI As Long
   LI = cmbChoMac.ListIndex
   HSCho(0).Value = 0
   HSCho(1).Value = &H40
   HSCho(2).Value = ChoM(LI).cFB
   HSCho(3).Value = ChoM(LI).cDL
   HSCho(4).Value = ChoM(LI).cRT
   HSCho(5).Value = ChoM(LI).cDL
   HSCho(6).Value = 0
   If DoSend = False Then Exit Sub
   dmpB(50) = CByte(LI)
   If Send = False Then Exit Sub
   Dim ComStr As String
   ComStr = Syx1 & Chr(&H1) & Chr(&H38) & Chr(LI)
   ComStr = ComStr & Chr(-(&H40 + &H1 + &H38 + LI) And 127)
   ComStr = ComStr & Chr(&HF7)
   SysExDT1 ComStr
End Sub

Private Sub cmbRevMac_Click()
   Dim LI As Long
   LI = cmbRevMac.ListIndex
   HSRev(0).Value = LI
   HSRev(1).Value = RevM(LI).rPL
   HSRev(2).Value = &H40
   HSRev(3).Value = RevM(LI).rTM
   HSRev(4).Value = RevM(LI).rFB
   HSRev(5).Value = 0
   If DoSend = False Then Exit Sub
   dmpB(42) = CByte(LI)
   If Send = False Then Exit Sub
   Dim ComStr As String
   ComStr = Syx1 & Chr(&H1) & Chr(&H30) & Chr(LI)
   ComStr = ComStr & Chr(-(&H40 + &H1 + &H30 + LI) And 127)
   ComStr = ComStr & Chr(&HF7)
   SysExDT1 ComStr
End Sub

Private Sub Form_Load()
   Dim I As Long
   
   Syx1 = Chr(&HF0) & Chr(&H41) & Chr(&H10) & Chr(&H42) & Chr(&H12) & Chr(&H40)
   If hMidiOUT <> 0 Then chkSend.Value = 1

   For I = 0 To 7: cmbRevMac.AddItem RevM(I).name: Next I
   For I = 0 To 7: cmbChoMac.AddItem ChoM(I).name: Next I
   InitALL
   DoSend = True
End Sub

Private Sub hsCho_Change(Index As Integer)
   Dim ComStr As String
   lblValCho(Index).Caption = HSCho(Index).Value
   If DoSend = False Then Exit Sub
   dmpB(51 + Index) = HSCho(Index).Value
   If Send = False Then Exit Sub
   ComStr = Syx1 & Chr(&H1) & Chr(&H39 + Index) & Chr(HSCho(Index).Value)
   ComStr = ComStr & Chr(-(&H40 + &H1 + &H39 + Index + HSCho(Index).Value) And 127)
   ComStr = ComStr & Chr(&HF7)
   SysExDT1 ComStr
End Sub

Private Sub hsCho_Scroll(Index As Integer)
   lblValCho(Index).Caption = HSCho(Index).Value
   If DoSend = False Then Exit Sub
   dmpB(51 + Index) = HSCho(Index).Value
   If Send = False Then Exit Sub
   Dim ComStr As String
   ComStr = Syx1 & Chr(&H1) & Chr(&H39 + Index) & Chr(HSCho(Index).Value)
   ComStr = ComStr & Chr(-(&H40 + &H1 + &H39 + Index + HSCho(Index).Value) And 127)
   ComStr = ComStr & Chr(&HF7)
   SysExDT1 ComStr
End Sub

Private Sub hsMas_Change(Index As Integer)
   If Index = 1 Then
      lblValMas(Index).Caption = isPanPot(HSMas(Index).Value)
      Else
      lblValMas(Index).Caption = HSMas(Index).Value
      End If
   
   If DoSend = False Then Exit Sub
   Select Case Index
      Case 0: dmpB(2) = HSMas(Index).Value
      Case 1: dmpB(6) = HSMas(Index).Value
      'Case 2: dmpB(0) = hsmas(Index).Value
      Case 3: dmpB(5) = 64 + HSMas(Index).Value
   End Select
   
   If Send = False Then Exit Sub
   Dim ComStr As String
   Dim jA As Byte, oB As Byte
   Select Case Index
      Case 0: jA = 4: oB = HSMas(Index).Value
      Case 1: jA = 6: oB = HSMas(Index).Value
      'Case 2: ja=0:hsmas(Index).Value
      Case 3: jA = 5: oB = 64 + HSMas(Index).Value
   End Select
   
   ComStr = Syx1 & Chr(&H0) & Chr(jA) & Chr(oB)
   ComStr = ComStr & Chr(-(&H40 + &H0 + jA + oB) And 127)
   ComStr = ComStr & Chr(&HF7)
   SysExDT1 ComStr
End Sub

Private Sub hsMas_Scroll(Index As Integer)
   If Index = 1 Then
      lblValMas(Index).Caption = isPanPot(HSMas(Index).Value)
      Else
      lblValMas(Index).Caption = HSMas(Index).Value
      End If

   If DoSend = False Then Exit Sub
   Select Case Index
      Case 0: dmpB(2) = HSMas(Index).Value
      Case 1: dmpB(6) = HSMas(Index).Value
      'Case 2: dmpB(0) = hsmas(Index).Value
      Case 3: dmpB(5) = 64 + HSMas(Index).Value
   End Select
   If Send = False Then Exit Sub
   Dim ComStr As String
   Dim jA As Byte, oB As Byte
   Select Case Index
      Case 0: jA = 4: oB = HSMas(Index).Value
      Case 1: jA = 6: oB = HSMas(Index).Value
      'Case 2: ja=0:hsmas(Index).Value
      Case 3: jA = 5: oB = 64 + HSMas(Index).Value
   End Select
   
   ComStr = Syx1 & Chr(&H0) & Chr(jA) & Chr(oB)
   ComStr = ComStr & Chr(-(&H40 + &H0 + jA + oB) And 127)
   ComStr = ComStr & Chr(&HF7)
   SysExDT1 ComStr
End Sub

Private Sub hsRev_Change(Index As Integer)
   lblValRev(Index).Caption = HSRev(Index).Value
   If DoSend = False Then Exit Sub
   dmpB(43 + Index) = HSRev(Index).Value
   If Send = False Then Exit Sub
   Dim ComStr As String
   ComStr = Syx1 & Chr(&H1) & Chr(&H31 + Index) & Chr(HSRev(Index).Value)
   ComStr = ComStr & Chr(-(&H40 + &H1 + &H31 + Index + HSRev(Index).Value) And 127)
   ComStr = ComStr & Chr(&HF7)
   SysExDT1 ComStr
End Sub

Private Sub hsRev_Scroll(Index As Integer)
   lblValRev(Index).Caption = HSRev(Index).Value
   If DoSend = False Then Exit Sub
   dmpB(43 + Index) = HSRev(Index).Value
   If Send = False Then Exit Sub
   Dim ComStr As String
   ComStr = Syx1 & Chr(&H1) & Chr(&H31 + Index) & Chr(HSRev(Index).Value)
   ComStr = ComStr & Chr(-(&H40 + &H1 + &H31 + Index + HSRev(Index).Value) And 127)
   ComStr = ComStr & Chr(&HF7)
   SysExDT1 ComStr
End Sub


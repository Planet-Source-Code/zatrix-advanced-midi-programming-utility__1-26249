VERSION 5.00
Begin VB.Form frmReadMid 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Read Midi File"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   ControlBox      =   0   'False
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      DrawMode        =   8  'Xor Pen
      FillColor       =   &H0000C0C0&
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
      ForeColor       =   &H00C00000&
      Height          =   345
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   4035
      TabIndex        =   6
      Top             =   570
      Visible         =   0   'False
      Width           =   4065
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   525
      Left            =   210
      TabIndex        =   2
      Top             =   90
      Width           =   5715
      Begin VB.CheckBox chkFilter 
         Caption         =   "Ctl Ch"
         Height          =   195
         Index           =   1
         Left            =   1890
         TabIndex        =   7
         Top             =   120
         Width           =   855
      End
      Begin VB.CheckBox chkFilter 
         Caption         =   "Syx Ex"
         Height          =   195
         Index           =   2
         Left            =   2940
         TabIndex        =   5
         Top             =   120
         Width           =   1065
      End
      Begin VB.CheckBox chkFilter 
         Caption         =   "Note msg"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   4
         Top             =   120
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.Label lbl 
         Caption         =   "Filter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   465
      Left            =   2970
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   465
      Left            =   1530
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "frmReadMid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' since reading midi files can take a while, and sometimes the
' resulting text is larger then 64kB, this form offers the
' possibility of Canceling the action, viewing the progress,
' filtering the midi messages
Option Explicit

Dim mMax As Long
Public Sub SetMax(ByVal NewMax As Long)
   If NewMax > 0 Then mMax = NewMax Else mMax = 100
   pic.Scale (0, 0)-(mMax, 1)
End Sub

Public Sub SetProgress(ByVal Pos As Long)
   Dim txt As String
   txt = CStr(Int(Pos / mMax * 100)) & " %"
   pic.Cls
   pic.CurrentX = (mMax - pic.TextWidth(txt)) \ 2
   pic.CurrentY = 0
   pic.Print txt
   pic.Line (0, 0)-(Pos, 1), pic.FillColor, BF
End Sub

Private Sub chkFilter_Click(Index As Integer)
   Select Case Index
      Case 0: FilterNoteMsg = IIf(chkFilter(Index).Value = 1, True, False)
      Case 1: FilterCtlChMsg = IIf(chkFilter(Index).Value = 1, True, False)
      Case 2: FilterSysExMsg = IIf(chkFilter(Index).Value = 1, True, False)
   End Select
End Sub

Private Sub cmdCancel_Click()
   Cancel = True
End Sub

Private Sub cmdOK_Click()
   cmdOK.Visible = False
   pic.Visible = True
   Frame1.Enabled = False
   DoEvents
   OK = True
End Sub

Private Sub Form_Load()
   FilterNoteMsg = True
   SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub


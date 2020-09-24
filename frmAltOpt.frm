VERSION 5.00
Begin VB.Form frmAltOpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alternatives"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   300
      Left            =   3510
      TabIndex        =   28
      Top             =   2280
      Width           =   405
   End
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   300
      Left            =   300
      TabIndex        =   27
      Text            =   "DumpAlt.mid"
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CheckBox chkNoCommon 
      Caption         =   "No Common"
      Height          =   195
      Left            =   270
      TabIndex        =   26
      Top             =   1200
      Width           =   1725
   End
   Begin VB.Frame Frame3 
      Caption         =   "Filter Parts"
      Height          =   1845
      Left            =   2310
      TabIndex        =   9
      Top             =   210
      Width           =   1605
      Begin VB.CheckBox chkFP 
         Caption         =   "16"
         Height          =   345
         Index           =   15
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1350
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "15"
         Height          =   345
         Index           =   14
         Left            =   810
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1350
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "14"
         Height          =   345
         Index           =   13
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1350
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "13"
         Height          =   345
         Index           =   12
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1350
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "12"
         Height          =   345
         Index           =   11
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   990
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "11"
         Height          =   345
         Index           =   10
         Left            =   810
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   990
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "10"
         Height          =   345
         Index           =   9
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   990
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "9"
         Height          =   345
         Index           =   8
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   990
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "8"
         Height          =   345
         Index           =   7
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   630
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "7"
         Height          =   345
         Index           =   6
         Left            =   810
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   630
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "6"
         Height          =   345
         Index           =   5
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   630
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "5"
         Height          =   345
         Index           =   4
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   630
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "4"
         Height          =   345
         Index           =   3
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   270
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "3"
         Height          =   345
         Index           =   2
         Left            =   810
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   270
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "2"
         Height          =   345
         Index           =   1
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   270
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "1"
         Height          =   345
         Index           =   0
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   270
         Width           =   315
      End
   End
   Begin VB.CheckBox chkReset 
      Caption         =   "Start with GS Reset"
      Height          =   195
      Left            =   270
      TabIndex        =   8
      Top             =   900
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   585
      Left            =   150
      TabIndex        =   5
      Top             =   1560
      Width           =   1965
      Begin VB.OptionButton optMake 
         Caption         =   "Save as  *.mid"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   330
         Width           =   2565
      End
      Begin VB.OptionButton optMake 
         Caption         =   "Just show "
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   30
         Value           =   -1  'True
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   705
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2025
      Begin VB.OptionButton optUse 
         Caption         =   "Sys Ex && Short Msg"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   390
         Value           =   -1  'True
         Width           =   2685
      End
      Begin VB.OptionButton optUse 
         Caption         =   "Sys Ex only"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   60
         Width           =   2325
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   465
      Left            =   450
      TabIndex        =   1
      Top             =   2910
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   465
      Left            =   2190
      TabIndex        =   0
      Top             =   2910
      Width           =   1575
   End
End
Attribute VB_Name = "frmAltOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form allows the user to set some option regarding to how
' the alternatives for the dump should be generated
Option Explicit

' exclude parts
Private Sub chkFP_Click(Index As Integer)
   PartSwitch(Index) = IIf(chkFP(Index).Value = 1, True, False)
End Sub

' exclude common/all
Private Sub chkNoCommon_Click()
   If chkNoCommon.Value = 1 Then
      chkReset.Value = 0
      chkReset.Enabled = False
      AltNoCommon = True
      Else
      chkReset.Enabled = True
      chkReset.Value = 1
      AltNoCommon = False
      End If
End Sub

' include GSreset
Private Sub chkReset_Click()
   StartWithGSReset = IIf(chkReset.Value = 1, True, False)
End Sub

' provide a filename to save the altern.
Private Sub cmdBrowse_Click()
   CD_File.hWndOwner = Me.hwnd
   CD_File.FileName = GetFileTitle(txtFile.text) & Chr(0)
   CD_File.filter = "Midi file |*.mid|All Files (*.*)|*.*"
   CD_File.Flags = OFN_HIDEREADONLY Or OFN_EXPLORER
   CD_File.MaxFileSize = 255
   On Error Resume Next
   CD_File.ShowOpen
   If err = 32755 Then
      Exit Sub
   ElseIf err = 0 Then
      txtFile.text = CD_File.FileName
   Else
      MsgBox err.Description
   End If
End Sub

Private Sub cmdCancel_Click()
   OK = False: Unload Me
End Sub

Private Sub cmdOK_Click()
   OK = True: Unload Me
End Sub

Private Sub Form_Load()
   Dim J As Integer
   ' set the checkboxes on/off
   For J = 0 To 15: chkFP(J).Value = IIf(PartSwitch(J), 1, 0): Next J
   SaveAlternat = False ' don't save just show
   UseShMsg = True ' use short messages when possible
   StartWithGSReset = True
   AltNoCommon = False
   SaveDumpAltFile = "DumpAlt.mid"
End Sub

Private Sub optUse_Click(Index As Integer)
   UseShMsg = IIf(Index = 0, False, True)
End Sub

Private Sub optMake_Click(Index As Integer)
   SaveAlternat = IIf(Index = 0, False, True)
End Sub

Private Sub txtFile_Change()
   SaveDumpAltFile = txtFile.text
End Sub


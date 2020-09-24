VERSION 5.00
Begin VB.Form frmMixLong 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SYX edit & save"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   Icon            =   "frmMixLong.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSend 
      Height          =   315
      Left            =   5940
      Picture         =   "frmMixLong.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Midi Out"
      Top             =   840
      UseMaskColor    =   -1  'True
      Value           =   2  'Grayed
      Width           =   345
   End
   Begin VB.Frame fraChan 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   2280
      TabIndex        =   22
      Top             =   1590
      Width           =   4605
      Begin VB.OptionButton optChan 
         Caption         =   "1"
         Height          =   330
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   0
         Value           =   -1  'True
         Width           =   270
      End
      Begin VB.OptionButton optChan 
         Caption         =   "2"
         Height          =   330
         Index           =   1
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   0
         Width           =   270
      End
      Begin VB.OptionButton optChan 
         Caption         =   "3"
         Height          =   330
         Index           =   2
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   0
         Width           =   270
      End
      Begin VB.OptionButton optChan 
         Caption         =   "4"
         Height          =   330
         Index           =   3
         Left            =   810
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   0
         Width           =   270
      End
      Begin VB.OptionButton optChan 
         Caption         =   "5"
         Height          =   330
         Index           =   4
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   0
         Width           =   270
      End
      Begin VB.OptionButton optChan 
         Caption         =   "6"
         Height          =   330
         Index           =   5
         Left            =   1350
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   0
         Width           =   270
      End
      Begin VB.OptionButton optChan 
         Caption         =   "7"
         Height          =   330
         Index           =   6
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   0
         Width           =   270
      End
      Begin VB.OptionButton optChan 
         Caption         =   "8"
         Height          =   330
         Index           =   7
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   0
         Width           =   270
      End
      Begin VB.OptionButton optChan 
         Caption         =   "9"
         Height          =   330
         Index           =   8
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   0
         Width           =   270
      End
      Begin VB.OptionButton optChan 
         Caption         =   "10"
         Height          =   330
         Index           =   9
         Left            =   2430
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   0
         Width           =   300
      End
      Begin VB.OptionButton optChan 
         Caption         =   "11"
         Height          =   330
         Index           =   10
         Left            =   2730
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   0
         Width           =   300
      End
      Begin VB.OptionButton optChan 
         Caption         =   "12"
         Height          =   330
         Index           =   11
         Left            =   3030
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   0
         Width           =   300
      End
      Begin VB.OptionButton optChan 
         Caption         =   "13"
         Height          =   330
         Index           =   12
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   0
         Width           =   300
      End
      Begin VB.OptionButton optChan 
         Caption         =   "14"
         Height          =   330
         Index           =   13
         Left            =   3630
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   0
         Width           =   300
      End
      Begin VB.OptionButton optChan 
         Caption         =   "15"
         Height          =   330
         Index           =   14
         Left            =   3930
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   0
         Width           =   300
      End
      Begin VB.OptionButton optChan 
         Caption         =   "16"
         Height          =   330
         Index           =   15
         Left            =   4230
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         Width           =   300
      End
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Def"
      Height          =   315
      Left            =   6300
      TabIndex        =   21
      ToolTipText     =   "Set to default value"
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmdSave 
      Height          =   345
      Index           =   2
      Left            =   2610
      Picture         =   "frmMixLong.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Prent bewaren"
      Top             =   3540
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdCopy 
      Height          =   345
      Index           =   2
      Left            =   2250
      Picture         =   "frmMixLong.frx":0506
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Geheel of selectie kopiëren"
      Top             =   3540
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdSave 
      Height          =   345
      Index           =   1
      Left            =   2610
      Picture         =   "frmMixLong.frx":0600
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Prent bewaren"
      Top             =   2865
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdCopy 
      Height          =   345
      Index           =   1
      Left            =   2250
      Picture         =   "frmMixLong.frx":06FA
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Geheel of selectie kopiëren"
      Top             =   2865
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdSave 
      Height          =   345
      Index           =   0
      Left            =   2610
      Picture         =   "frmMixLong.frx":07F4
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Prent bewaren"
      Top             =   2220
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdCopy 
      Height          =   345
      Index           =   0
      Left            =   2250
      Picture         =   "frmMixLong.frx":08EE
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Geheel of selectie kopiëren"
      Top             =   2220
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.ListBox lstParam 
      Height          =   3765
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1995
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2550
      ScaleHeight     =   285
      ScaleWidth      =   3105
      TabIndex        =   0
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label lbl 
      Caption         =   "ASCII-code (syx)"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   2250
      TabIndex        =   20
      Top             =   3300
      Width           =   1365
   End
   Begin VB.Label lbl 
      Caption         =   "Decimal"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   2250
      TabIndex        =   19
      Top             =   2625
      Width           =   1125
   End
   Begin VB.Label lbl 
      Caption         =   "Hexadecimal"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   2250
      TabIndex        =   18
      Top             =   1980
      Width           =   1125
   End
   Begin VB.Label lblSyx 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblSyx"
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   2
      Left            =   3000
      TabIndex        =   17
      Top             =   3540
      Width           =   3825
   End
   Begin VB.Label lblSyx 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblSyx"
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   1
      Left            =   3000
      TabIndex        =   14
      Top             =   2865
      Width           =   3825
   End
   Begin VB.Label lblSyx 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblSyx"
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   0
      Left            =   3000
      TabIndex        =   11
      Top             =   2220
      Width           =   3825
   End
   Begin VB.Label lblCurPar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblCurPar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   345
      Left            =   2250
      TabIndex        =   8
      Top             =   120
      Width           =   4545
   End
   Begin VB.Label lblDec 
      Alignment       =   2  'Center
      Caption         =   "0000000"
      Height          =   195
      Index           =   2
      Left            =   5340
      TabIndex        =   6
      Top             =   1200
      Width           =   675
   End
   Begin VB.Label lblDec 
      Alignment       =   2  'Center
      Caption         =   "0000000"
      Height          =   195
      Index           =   1
      Left            =   3780
      TabIndex        =   5
      Top             =   1200
      Width           =   675
   End
   Begin VB.Label lblDec 
      Alignment       =   2  'Center
      Caption         =   "0000000"
      Height          =   195
      Index           =   0
      Left            =   2250
      TabIndex        =   4
      Top             =   1200
      Width           =   675
   End
   Begin VB.Label lblHex 
      Alignment       =   2  'Center
      Caption         =   "0000"
      Height          =   195
      Index           =   2
      Left            =   5490
      TabIndex        =   3
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblHex 
      Alignment       =   2  'Center
      Caption         =   "0000"
      Height          =   195
      Index           =   1
      Left            =   3930
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblHex 
      Alignment       =   2  'Center
      Caption         =   "0000"
      Height          =   195
      Index           =   0
      Left            =   2370
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
End
Attribute VB_Name = "frmMixLong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form offers the possibility to generate *.SYX files for
' separate individual system exclusive messages.
Option Explicit
Dim lstOffs As Long     ' offset to parts param in lstParam
Dim CurParam As Long    ' curren param index (<>listindex)
Dim CurGroup As Long    ' common(0) or parts(1)
Dim CurValue As Long    ' current value of current param
Dim CurType As Byte     ' var type param
Dim CurComStr As String ' SysEx command string
Dim Send As Boolean     ' send changes to sound canvas

' get Common/All parameter value
Function GetAParam(ByVal ParamID As Long) As Long
   Dim s As Long
   If ParamID = 0 Then
      s = dmpB(0) + dmpB(1) * 256 ' master volume 2 bytes
      Else
      s = dmpB(APar(ParamID).ByteOffs)
      End If
   GetAParam = s
End Function

' get Parts parameter value (Channel~Part)
Function GetPParam(ByVal CHANNEL As Long, ByVal ParamID As Long) As Long
   Dim B As Byte
   Dim s As Long
   If PPar(ParamID).ShortName = "PIT FIN" Then
      s = dmpB(72 + CHANNEL * 112 + PPar(ParamID).ByteOffs)
      s = s + 256 * dmpB(72 + CHANNEL * 112 + PPar(ParamID).ByteOffs + 1)
      Else
      s = dmpB(72 + CHANNEL * 112 + PPar(ParamID).ByteOffs)
      End If
   GetPParam = s
End Function

Private Sub Init()
   Dim I As Long  ' counter
   Dim C As Long  ' count parts param oofset
   
   For I = 1 To AParCount - 1
      ' excluding  4:patchname  5:part.res.  7:rev char.
      If Not (I = 4 Or I = 5 Or I = 7) Then
         lstParam.AddItem Trim(APar(I).name)
         lstParam.ItemData(lstParam.NewIndex) = I
         C = C + 1
         End If
   Next I
   lstOffs = C
   For I = 0 To PParCount - 1
      ' excluding  2:channel  3:M/P drum  36:Rx.1   37:Rx.2
      If Not (I = 2 Or I = 3 Or I = 36 Or I = 37) Then
         lstParam.AddItem Trim(PPar(I).name)
         lstParam.ItemData(lstParam.NewIndex) = I
         End If
   Next I
   lstParam.ListIndex = 0
End Sub

' set All parameter value
Private Sub SetAParam(ByVal ParamID As Long, ByVal Value As Byte)
   If ParamID = 0 Then
      dmpB(APar(ParamID).ByteOffs) = Value And 255
      dmpB(APar(ParamID).ByteOffs + 1) = (Value \ 256) And 255
      Else
      dmpB(APar(ParamID).ByteOffs) = Value
      End If
End Sub

' set parts parameter value
Private Sub SetPParam(ByVal CHANNEL As Long, ByVal ParamID As Long, ByVal Value As Long)
   If PPar(ParamID).ShortName = "PIT FIN" Then
      dmpB(72 + CHANNEL * 112 + PPar(ParamID).ByteOffs) = Value And 255
      dmpB(72 + CHANNEL * 112 + PPar(ParamID).ByteOffs + 1) = (Value \ 256) And 255
      Else
      dmpB(72 + CHANNEL * 112 + PPar(ParamID).ByteOffs) = Value
      End If
End Sub

' when a different param is selected
Private Sub ShowCurParam()
   Dim minH As Single, maxH As Single
   Dim minD As Single, maxD As Single
   
   CurParam = lstParam.ItemData(lstParam.ListIndex)
   If lstParam.ListIndex < lstOffs Then
      CurGroup = 0 'apar
      minH = PType(APar(CurParam).Type).MinHex
      maxH = PType(APar(CurParam).Type).MaxHex
      minD = PType(APar(CurParam).Type).MinDec
      maxD = PType(APar(CurParam).Type).MaxDec
      CurValue = GetAParam(CurParam)
      CurType = APar(CurParam).Type
      lblCurPar.Caption = APar(CurParam).name
      fraChan.Visible = False
      Else
      CurGroup = 1 'ppar
      minH = PType(PPar(CurParam).Type).MinHex
      maxH = PType(PPar(CurParam).Type).MaxHex
      minD = PType(PPar(CurParam).Type).MinDec
      maxD = PType(PPar(CurParam).Type).MaxDec
      CurValue = GetPParam(CurDChannel, CurParam)
      CurType = PPar(CurParam).Type
      lblCurPar.Caption = PPar(CurParam).name
      fraChan.Visible = True
      End If
   lblHex(0).Caption = Hex(minH)
   lblHex(2).Caption = Hex(maxH)
   lblDec(0).Caption = minD
   lblDec(2).Caption = maxD
   pic.Scale (minH, 0)-(maxH, 1) ' adapt picture scale to param type
   ShowCurValue
End Sub

' when a new value is selected
Private Sub ShowCurValue()
   lblHex(1).Caption = Hex(CurValue)
   lblDec(1).Caption = isValue(Hex(CurValue), CurType)
   pic.Cls
   pic.Line (0, 0)-(CurValue, 1), QBColor(9), BF
   If CurGroup = 0 Then
      CurComStr = makeComStr(APar(CurParam).Address, CurDChannel, CurValue, CurType, False)
      Else
      CurComStr = makeComStr(PPar(CurParam).Address, CurDChannel, CurValue, CurType, False)
      SetPParam CurDChannel, CurParam, CurValue
      End If
   lblSyx(0).Caption = getComStrHex(CurComStr)
   lblSyx(1).Caption = getComStrDec(CurComStr)
   lblSyx(2).Caption = getComStrStr(CurComStr)
   If Send = False Then Exit Sub
   SysExDT1 CurComStr
End Sub

' send to sound canvas
Private Sub chkSend_Click()
   If chkSend.Value = 1 Then
      If hMidiOUT = 0 Then MidiOUT_Port "open"
      If hMidiOUT = 0 Then chkSend.Value = 0
      Send = True
      Else
      Send = False
      End If
End Sub

Private Sub cmdCopy_Click(Index As Integer)
   Clipboard.Clear
   If Index < 2 Then
      Clipboard.SetText lblSyx(Index).Caption ' hex/dec
      Else
      Clipboard.SetText CurComStr ' byte string
      End If
End Sub

' set current param value to default
Private Sub cmdDefault_Click()
   If CurGroup = 0 Then
      CurValue = APar(CurParam).Default
      Else
      CurValue = PPar(CurParam).Default
      End If
   ShowCurValue
End Sub

Private Sub cmdSave_Click(Index As Integer)
   Dim ch As Long
   Dim filter As String, def As String
   
   If Index = 2 Then
      filter = "Sys Ex file |*.syx|All Files (*.*)|*.*"
      def = ".syx"
      Else
      filter = "Text file |*.txt|All Files (*.*)|*.*"
      def = ".txt"
      End If
   
   CD_File.hWndOwner = Me.hwnd
   CD_File.FileName = Chr(0)
   CD_File.filter = filter
   CD_File.Flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
   CD_File.MaxFileSize = 255
   CD_File.DefaultExt = def
   On Error Resume Next
   CD_File.ShowSave
   If err = 32755 Then
      Exit Sub
   ElseIf err = 0 Then
      Screen.MousePointer = 11
      ch = FreeFile
      Open CD_File.FileName For Output As ch
      If Index < 2 Then
         Print #ch, lblSyx(Index).Caption
         Else
         Print #ch, CurComStr;
         End If
      Close ch
      If err <> 0 Then MsgBox err.Description
      Screen.MousePointer = 0
   Else
      MsgBox err.Description
   End If

End Sub

Private Sub Form_Activate()
   optChan(CurChannel).Value = True
End Sub

Private Sub Form_Load()
   Init
'   SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub lstParam_Click()
   ShowCurParam
End Sub

Private Sub optChan_Click(Index As Integer)
   SetChannel Index
   CurValue = GetPParam(CurDChannel, CurParam)
   ShowCurValue
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   CurValue = X
   ShowCurValue
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 0 Then Exit Sub
   If X >= pic.ScaleLeft And X <= pic.ScaleLeft + pic.ScaleWidth Then
      CurValue = X
      ShowCurValue
      End If
End Sub


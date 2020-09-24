VERSION 5.00
Begin VB.Form frmSeq 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fade Sequence"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   317
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   300
      Left            =   1320
      TabIndex        =   61
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame FraAll 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4695
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6525
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   300
         Left            =   5880
         TabIndex        =   60
         Top             =   3420
         Width           =   405
      End
      Begin VB.TextBox txtFile 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2670
         TabIndex        =   59
         Text            =   "FadeOut.mid"
         Top             =   3420
         Width           =   3135
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   585
         Left            =   2880
         TabIndex        =   40
         Top             =   3900
         Width           =   1575
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Save"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   4710
         TabIndex        =   39
         Top             =   3900
         Width           =   1575
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "16"
         Height          =   345
         Index           =   15
         Left            =   5970
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   30
         Value           =   1  'Checked
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "15"
         Height          =   345
         Index           =   14
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   30
         Value           =   1  'Checked
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "14"
         Height          =   345
         Index           =   13
         Left            =   5190
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   30
         Value           =   1  'Checked
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "13"
         Height          =   345
         Index           =   12
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   30
         Value           =   1  'Checked
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "12"
         Height          =   345
         Index           =   11
         Left            =   4410
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   30
         Value           =   1  'Checked
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "11"
         Height          =   345
         Index           =   10
         Left            =   4020
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   30
         Value           =   1  'Checked
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "10"
         Height          =   345
         Index           =   9
         Left            =   3630
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   30
         Value           =   1  'Checked
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "9"
         Height          =   345
         Index           =   8
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   30
         Value           =   1  'Checked
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "8"
         Height          =   345
         Index           =   7
         Left            =   2850
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   30
         Value           =   1  'Checked
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "7"
         Height          =   345
         Index           =   6
         Left            =   2460
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   30
         Value           =   1  'Checked
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "6"
         Height          =   345
         Index           =   5
         Left            =   2070
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   30
         Value           =   1  'Checked
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "5"
         Height          =   345
         Index           =   4
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   30
         Value           =   1  'Checked
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "4"
         Height          =   345
         Index           =   3
         Left            =   1290
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   30
         Value           =   1  'Checked
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "3"
         Height          =   345
         Index           =   2
         Left            =   900
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   30
         Value           =   1  'Checked
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "2"
         Height          =   345
         Index           =   1
         Left            =   510
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   30
         Value           =   1  'Checked
         Width           =   315
      End
      Begin VB.CheckBox chkFP 
         Caption         =   "1"
         Height          =   345
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   30
         Value           =   1  'Checked
         Width           =   315
      End
      Begin VB.PictureBox picV 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1950
         Index           =   0
         Left            =   120
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   22
         Top             =   780
         Width           =   315
      End
      Begin VB.PictureBox picV 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1950
         Index           =   1
         Left            =   510
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   21
         Top             =   780
         Width           =   315
      End
      Begin VB.PictureBox picV 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1950
         Index           =   2
         Left            =   900
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   20
         Top             =   780
         Width           =   315
      End
      Begin VB.PictureBox picV 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1950
         Index           =   3
         Left            =   1290
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   19
         Top             =   780
         Width           =   315
      End
      Begin VB.PictureBox picV 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1950
         Index           =   4
         Left            =   1680
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   18
         Top             =   780
         Width           =   315
      End
      Begin VB.PictureBox picV 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1950
         Index           =   5
         Left            =   2070
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   17
         Top             =   780
         Width           =   315
      End
      Begin VB.PictureBox picV 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1950
         Index           =   6
         Left            =   2460
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   16
         Top             =   780
         Width           =   315
      End
      Begin VB.PictureBox picV 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1950
         Index           =   7
         Left            =   2850
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   15
         Top             =   780
         Width           =   315
      End
      Begin VB.PictureBox picV 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1950
         Index           =   8
         Left            =   3240
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   14
         Top             =   780
         Width           =   315
      End
      Begin VB.PictureBox picV 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1950
         Index           =   9
         Left            =   3630
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   13
         Top             =   780
         Width           =   315
      End
      Begin VB.PictureBox picV 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1950
         Index           =   10
         Left            =   4020
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   12
         Top             =   780
         Width           =   315
      End
      Begin VB.PictureBox picV 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1950
         Index           =   11
         Left            =   4410
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   11
         Top             =   780
         Width           =   315
      End
      Begin VB.PictureBox picV 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1950
         Index           =   12
         Left            =   4800
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   10
         Top             =   780
         Width           =   315
      End
      Begin VB.PictureBox picV 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1950
         Index           =   13
         Left            =   5190
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   9
         Top             =   780
         Width           =   315
      End
      Begin VB.PictureBox picV 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1950
         Index           =   14
         Left            =   5580
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   8
         Top             =   780
         Width           =   315
      End
      Begin VB.PictureBox picV 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1950
         Index           =   15
         Left            =   5970
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   7
         Top             =   780
         Width           =   315
      End
      Begin VB.CommandButton cmdDumpVal 
         Caption         =   "Dump values"
         Height          =   585
         Left            =   120
         TabIndex        =   6
         Top             =   3900
         Width           =   1215
      End
      Begin VB.TextBox txtSteps 
         Height          =   300
         Left            =   3270
         TabIndex        =   5
         Text            =   "10"
         ToolTipText     =   "Number of steps"
         Top             =   2850
         Width           =   585
      End
      Begin VB.TextBox txtMin 
         Height          =   300
         Left            =   4380
         TabIndex        =   4
         Text            =   "32"
         ToolTipText     =   "Minimum volume"
         Top             =   2850
         Width           =   585
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   435
         Left            =   5130
         TabIndex        =   1
         Top             =   2850
         Width           =   1185
         Begin VB.OptionButton optInOut 
            Caption         =   "OUT"
            Height          =   300
            Index           =   0
            Left            =   30
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Fade OUT"
            Top             =   0
            Value           =   -1  'True
            Width           =   555
         End
         Begin VB.OptionButton optInOut 
            Caption         =   "IN"
            Height          =   300
            Index           =   1
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Fade IN"
            Top             =   0
            Width           =   555
         End
      End
      Begin VB.Label lblV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   58
         Top             =   480
         Width           =   315
      End
      Begin VB.Label lblV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   1
         Left            =   510
         TabIndex        =   57
         Top             =   480
         Width           =   315
      End
      Begin VB.Label lblV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   2
         Left            =   900
         TabIndex        =   56
         Top             =   480
         Width           =   315
      End
      Begin VB.Label lblV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   3
         Left            =   1290
         TabIndex        =   55
         Top             =   480
         Width           =   315
      End
      Begin VB.Label lblV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   4
         Left            =   1680
         TabIndex        =   54
         Top             =   480
         Width           =   315
      End
      Begin VB.Label lblV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   5
         Left            =   2070
         TabIndex        =   53
         Top             =   480
         Width           =   315
      End
      Begin VB.Label lblV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   6
         Left            =   2460
         TabIndex        =   52
         Top             =   480
         Width           =   315
      End
      Begin VB.Label lblV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   7
         Left            =   2850
         TabIndex        =   51
         Top             =   480
         Width           =   315
      End
      Begin VB.Label lblV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   8
         Left            =   3240
         TabIndex        =   50
         Top             =   480
         Width           =   315
      End
      Begin VB.Label lblV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   9
         Left            =   3630
         TabIndex        =   49
         Top             =   480
         Width           =   315
      End
      Begin VB.Label lblV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   10
         Left            =   4020
         TabIndex        =   48
         Top             =   480
         Width           =   315
      End
      Begin VB.Label lblV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   11
         Left            =   4410
         TabIndex        =   47
         Top             =   480
         Width           =   315
      End
      Begin VB.Label lblV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   12
         Left            =   4800
         TabIndex        =   46
         Top             =   480
         Width           =   315
      End
      Begin VB.Label lblV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   13
         Left            =   5190
         TabIndex        =   45
         Top             =   480
         Width           =   315
      End
      Begin VB.Label lblV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   14
         Left            =   5580
         TabIndex        =   44
         Top             =   480
         Width           =   315
      End
      Begin VB.Label lblV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   15
         Left            =   5970
         TabIndex        =   43
         Top             =   480
         Width           =   315
      End
      Begin VB.Label lbl 
         Caption         =   "Steps"
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   42
         Top             =   2850
         Width           =   435
      End
      Begin VB.Label lbl 
         Caption         =   "Min"
         Height          =   195
         Index           =   1
         Left            =   4020
         TabIndex        =   41
         Top             =   2850
         Width           =   315
      End
   End
End
Attribute VB_Name = "frmSeq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form generates fade in and outs in the form of *.mid files
Option Explicit

Dim vol(16) As Byte

Private Sub Init()
   Dim J As Long
   Dim mChan As Long
   For J = 0 To 15
      mChan = isChannel(J)
      vol(J) = dmpB(72 + mChan * 112 + PPar(6).ByteOffs)
      picV(J).Line (0, 127 - vol(J))-(21, 127), QBColor(9), BF
      lblV(J).Caption = vol(J)
   Next J
End Sub

Private Sub SaveFade()
   Dim ch As Long
   Dim I As Long, J As Long, Lng As Long
   Dim txt As String, File As String, Doing As String
   Dim B As Byte, iB As Byte
   Dim DT As String * 1
   Dim Steps As Long
   Dim vMin As Long
   
   DT = Chr(24)
   Steps = Val(txtSteps.text) - 1
   vMin = Val(txtMin.text)
   If optInOut(0).Value = True Then
      Doing = "FadeOut"
      For I = 0 To Steps
         For J = 0 To 15
            If chkFP(J).Value = 1 Then
               iB = CByte(vol(J) - ((vol(J) - vMin) / Steps) * I)
               txt = txt & DT & makeAltShMsg(4, J, iB)
               End If
         Next J
      Next I
      Else
      Doing = "FadeIn"
      For I = 0 To Steps
         For J = 0 To 15
            If chkFP(J).Value = 1 Then
               iB = CByte(vMin + ((vol(J) - vMin) / Steps) * I)
               txt = txt & DT & makeAltShMsg(4, J, iB)
               End If
         Next J
      Next I
      End If

   File = SaveFadeFile
   If InStr(File, ":\") = 0 Then File = App.Path & "\" & File
   ch = FreeFile
   Open File For Binary As ch
   Put #ch, 1, "MThd"
   B = 0: Put #ch, , B: Put #ch, , B: Put #ch, , B: Put #ch, , CByte(6)
   Put #ch, , B: Put #ch, , B   ' formattype=0
   Put #ch, , B: Put #ch, , CByte(1)   ' numtracks
   Put #ch, , CByte(1): Put #ch, , CByte(128) ' division
   Put #ch, , "MTrk"
   Lng = 3 + Len(Doing) + Len(txt) + 3
   Put #ch, , CByte(Lng \ (256 ^ 3))
   Put #ch, , CByte((Lng \ (256 ^ 2)) And 255)
   Put #ch, , CByte((Lng \ 256) And 255)
   Put #ch, , CByte(Lng And 255)
   B = 0: Put #ch, , B
   B = &HFF: Put #ch, , B
   B = 3: Put #ch, , B
   B = Len(Doing): Put #ch, , B: Put #ch, , Doing
   Put #ch, , txt
   Put #ch, , DT
   B = &HFF: Put #ch, , B
   B = &H2F: Put #ch, , B
   B = 0: Put #ch, , B
   Close ch

End Sub

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
   Unload Me
End Sub

Private Sub cmdDumpVal_Click()
   Init
End Sub

Private Sub cmdOK_Click()
   Screen.MousePointer = vbHourglass
   SaveFade
   Screen.MousePointer = vbDefault
   Unload Me
End Sub

Private Sub cmdShow_Click()
   Dim I As Long, J As Long
   Dim Y As Long
   Dim Steps As Long
   Dim vMin As Long
   
   Screen.MousePointer = vbHourglass
   FraAll.Enabled = False
   Steps = Val(txtSteps.text) - 1
   vMin = Val(txtMin.text)
   If optInOut(0).Value = True Then
      For I = 0 To Steps
         For J = 0 To 15
            If chkFP(J).Value = 1 Then
               Y = vol(J) - ((vol(J) - vMin) / Steps) * I
               picV(J).Cls
               picV(J).Line (0, 127 - Y)-(21, 127), QBColor(9), BF
               lblV(J).Caption = Y
               End If
         Next J
         Pauze 500
      Next I
      Else
      For I = 0 To Steps
         For J = 0 To 15
            If chkFP(J).Value = 1 Then
               Y = vMin + ((vol(J) - vMin) / Steps) * I
               picV(J).Cls
               picV(J).Line (0, 127 - Y)-(21, 127), QBColor(9), BF
               lblV(J).Caption = Y
               End If
         Next J
         Pauze 500
      Next I
      End If
   Pauze 1000
   For J = 0 To 15
      picV(J).Cls
      picV(J).Line (0, 127 - vol(J))-(21, 127), QBColor(9), BF
      lblV(J).Caption = vol(J)
   Next J
   FraAll.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   Dim J As Long
   Init
   SaveFadeFile = "FadeOut.mid"
   For J = 0 To 15: chkFP(J).Value = IIf(PartSwitch(J), 0, 1): Next J
End Sub

Private Sub optInOut_Click(Index As Integer)
   If Index = 0 Then
      If txtFile.text = "FadeOut.mid" Then txtFile.text = "FadeIn.mid"
      Else
      If txtFile.text = "FadeIn.mid" Then txtFile.text = "FadeOut.mid"
      End If
End Sub

Private Sub picV_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   vol(Index) = 127 - Y
   lblV(Index) = vol(Index)
   picV(Index).Cls
   picV(Index).Line (0, Y)-(21, 127), QBColor(9), BF
End Sub

Private Sub picV_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 0 Then Exit Sub
   vol(Index) = 127 - Y
   lblV(Index) = vol(Index)
   picV(Index).Cls
   picV(Index).Line (0, Y)-(21, 127), QBColor(9), BF
End Sub

Private Sub txtFile_Change()
   SaveFadeFile = txtFile.text
End Sub


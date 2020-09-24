VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmBD 
   Caption         =   "Analyse SC55 bulk-dumps"
   ClientHeight    =   5205
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7935
   Icon            =   "frmBD.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox kadO 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   30
      ScaleHeight     =   1185
      ScaleWidth      =   2295
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3900
      Visible         =   0   'False
      Width           =   2355
      Begin VB.HScrollBar HScroll 
         Height          =   270
         LargeChange     =   1680
         Left            =   0
         Max             =   16890
         Min             =   -45
         SmallChange     =   240
         TabIndex        =   4
         Top             =   900
         Width           =   1050
      End
      Begin VB.PictureBox picKlav 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
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
         ForeColor       =   &H00FF00FF&
         Height          =   850
         Left            =   30
         MousePointer    =   2  'Cross
         ScaleHeight     =   57
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1120
         TabIndex        =   2
         Top             =   30
         Width           =   16800
         Begin VB.Label lblKey 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "000"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Visible         =   0   'False
            Width           =   435
         End
      End
   End
   Begin SHDocVwCtl.WebBrowser wbr 
      Height          =   3615
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5985
      ExtentX         =   10557
      ExtentY         =   6376
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileMidi 
         Caption         =   "&Midi"
         Begin VB.Menu mnuFileMidiRead 
            Caption         =   "&Read"
         End
         Begin VB.Menu mnuFileMidiPlay 
            Caption         =   "&Play"
         End
      End
      Begin VB.Menu mnuFileDump 
         Caption         =   "&Dump"
         Begin VB.Menu mnuFileDumpOpen 
            Caption         =   "&Open"
         End
         Begin VB.Menu mnuFileDumpSave 
            Caption         =   "&Save"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuFileDumpCompare 
            Caption         =   "&Compare"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuFileHTML 
         Caption         =   "&HTML"
         Begin VB.Menu mnuFileHTMLOpen 
            Caption         =   "&Open"
         End
         Begin VB.Menu mnuFileHTMLSave 
            Caption         =   "&Save"
         End
      End
      Begin VB.Menu mnuFilePart 
         Caption         =   "&Part"
         Begin VB.Menu mnuFilePartOpen 
            Caption         =   "&Open"
         End
         Begin VB.Menu mnuFilePartSave 
            Caption         =   "&Save"
         End
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuMidi 
      Caption         =   "&Midi"
      Begin VB.Menu mnuMidiDevices 
         Caption         =   "&Devices"
      End
      Begin VB.Menu mnuMidiSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMidiPortsOff 
         Caption         =   "&Close all ports"
      End
      Begin VB.Menu mnuMidiSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMidiThru 
         Caption         =   "&Midi Thru"
      End
      Begin VB.Menu mnuMidiKlavier 
         Caption         =   "&Keyboard"
      End
      Begin VB.Menu mnuMidiSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMix 
         Caption         =   "Mi&x"
         Begin VB.Menu mnuMixAll 
            Caption         =   "&All"
         End
         Begin VB.Menu mnuMixParts 
            Caption         =   "&Parts"
         End
      End
      Begin VB.Menu mnuMidiSeq 
         Caption         =   "&Fade in/out seq"
      End
      Begin VB.Menu mnuSysExSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMidiSYX 
         Caption         =   "&Micro edit && save"
      End
      Begin VB.Menu mnuMidiGSReset 
         Caption         =   "GS Reset"
      End
   End
   Begin VB.Menu mnuDump 
      Caption         =   "&Dump"
      Begin VB.Menu mnuDumpHTM 
         Caption         =   "&HTML"
         Begin VB.Menu mnuDumpHTML 
            Caption         =   "&Hex"
            Index           =   0
         End
         Begin VB.Menu mnuDumpHTML 
            Caption         =   "&Dec"
            Index           =   1
         End
      End
      Begin VB.Menu mnuDumpPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuDumpAlternat 
         Caption         =   "&Alternat"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "&Info"
      End
   End
End
Attribute VB_Name = "frmBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this is the main startup form
Option Explicit
Dim DocComplete As Boolean    ' html document load is completed
Dim CurKeyID As Long          ' remember note on for note off

' the data are presented here in the order they appear
' in the dump
Private Sub DumpCompare(ByVal File As String)
   Dim ch As Long, P As Long, I As Long, J As Long
   Dim offs As Long        ' syx starts at
   Dim inSYX As Boolean    ' in syx section of midi
   Dim preSYX As Long      ' syx headerbytes counter
   Dim MaxpreSYX As Long   ' max syx headerbytes
   Dim iB As Byte          ' IN byte
   Dim nB(7360) As Byte    ' nibblized bytes
   Dim nBc As Long         ' nibble bytes count
   Dim dB(3680) As Byte, dBc As Long
   Dim txt As String
   Dim pName As String
   
   ch = FreeFile
   Open File For Binary As ch
   ' find offset to syx
   For P = 1 To FileLen(File)
      Get #ch, P, iB
      If iB = &HF0 Then Exit For
   Next P
   offs = P
     
   inSYX = True: preSYX = 0: MaxpreSYX = 10
   nBc = 0
   For I = 1 To FileLen(File) - offs
      Get #ch, offs + I, nB(nBc)
      If nB(nBc) = &HF0 Then
         inSYX = True
         preSYX = 0 ' syx header counter reset
         End If
      If nB(nBc) = &HF7 Then
         inSYX = False
         nBc = nBc - 1 ' exclude checksum
         End If
      If inSYX = True Then
         If preSYX < MaxpreSYX Then
            If preSYX = 1 Then
               If nB(nBc) = &H81 Then ' packets of 128 bytes (&H81 Big Endian value)
                  MaxpreSYX = 10
                  Else
                  MaxpreSYX = 9 ' or smaller packet (2x)
                  End If
               End If
            preSYX = preSYX + 1
            Else
            nBc = nBc + 1
            If nBc = UBound(nB) Then Exit For
            End If
         End If
   Next I
   Close ch
   
   ' un-nibblize
   For I = 0 To nBc - 1 Step 2
      dB(I \ 2) = nB(I) * 16 Or nB(I + 1)
   Next I
   dBc = nBc \ 2

   txt = txt & "Packets seq. - SC55 dump compare: " & CurDmpMidFileTitle & "<>" & GetFileTitle(File) & vbCrLf & vbCrLf
   
   txt = txt & "Totaal aantal bytes: " & dmpBc & "<>" & dBc & vbCrLf & vbCrLf
   
   For I = 0 To 63
      pName = getParName(0, I)
      If I >= 8 And I <= 23 Then
         If I = 8 Then txt = txt & Format(I, "000") & " " & pName & " "
         If dmpB(I) = dB(I) Then txt = txt & Chr(dB(I)) Else txt = txt & "<B>" & Chr(dB(I)) & "</B> "
         If I = 23 Then txt = txt & vbCrLf
      ElseIf I >= 24 And I <= 39 Then
         If I = 24 Then txt = txt & Format(I, "000") & " " & pName & " "
         If dmpB(I) = dB(I) Then txt = txt & HexByte(dB(I)) & " " Else txt = txt & "<B>" & HexByte(dB(I)) & "</B> "
         If I = 39 Then txt = txt & vbCrLf
      Else
         If pName <> "" Then
            txt = txt & Format(I, "000") & " " & pName & " "
            If dmpB(I) = dB(I) Then txt = txt & HexByte(dB(I)) Else txt = txt & "<B>" & HexByte(dB(I)) & "</B> "
            txt = txt & vbCrLf
            End If
      End If
   Next I
   txt = txt & vbCrLf
   txt = txt & Space(25) & "10 01 02 03 04 05 06 07 08 09 11 12 13 14 15 16" & vbCrLf
   txt = txt & vbCrLf

   For I = 0 To 111
      pName = getParName(1, I)
      If pName <> "" Then
         txt = txt & Format(I, "000") & " " & getParName(1, I) & " "
         For J = 0 To 15
            iB = dB(72 + J * 112 + I)
            If dmpB(72 + J * 112 + I) = iB Then txt = txt & HexByte(iB) & " " Else txt = txt & "<B>" & HexByte(iB) & "</B> "
         Next J
         txt = txt & vbCrLf
         End If
   Next I
   SetDocument txt, True

End Sub

' routine for DumpCompare function
Private Function getParName(ByVal Group As Long, ByVal BO As Long) As String
   Dim I As Long
   If Group = 0 Then
      For I = 0 To AParCount - 1
         If APar(I).ByteOffs = BO Then getParName = APar(I).name: Exit For
      Next I
      Else
      For I = 0 To PParCount - 1
         If PPar(I).ByteOffs = BO Then getParName = PPar(I).name: Exit For
      Next I
      End If
End Function

Private Sub MakePiano(pic As PictureBox)
   Dim wX1 As Long, wY1 As Long  ' white key
   Dim wdX As Long, wdY As Long
   Dim zX1 As Long, zY1 As Long  ' black key
   Dim zdX As Long, zdY As Long
   Dim AaWTs As Long             ' white key count
   Dim I As Long                 ' counter
   
   wX1 = 0:  wY1 = 0: wdX = 16: wdY = 57 ' white key
   zX1 = 10: zY1 = 0: zdX = 11: zdY = 42 ' black key
   AaWTs = 11 * 7 ' 11 octaves with 7 white keys

   pic.Width = AaWTs * wdX * 15
   pic.AutoRedraw = True
   
   ' 1st white key
   pic.Line (wX1, wY1)-Step(wdX, wdY - 1), QBColor(15), BF
   BevelPic pic, wX1, wY1, wdX, wdY, False
   
   ' copy other white keys
   For I = 0 To AaWTs - 1
      BitBlt pic.hDC, wX1 + I * wdX, wY1, wdX, wdY, pic.hDC, wX1, wY1, SRCCOPY
   Next I
      
   ' 1st black key
   pic.Line (zX1, zY1)-Step(zdX, zdY), QBColor(0), BF
   BevelPic pic, zX1, zY1, zX1 + zdX, zdY, False
   
   ' copy other black key
   For I = 1 To AaWTs - 1
      If Mid("110111", (I Mod 7) + 1, 1) = "1" Then
         BitBlt pic.hDC, zX1 + I * wdX, zY1, zdX, zdY, pic.hDC, zX1, zY1, SRCCOPY
         End If
   Next I
   
   pic.Picture = pic.Image
   pic.AutoRedraw = False
End Sub

Public Sub DoNavigate(ByVal URL As String)
   DocComplete = False
   wbr.Navigate URL
   While DocComplete = False: DoEvents: Wend
End Sub


' browsers can read a simple text file, but here I mark non-default values
' so, the text isn't plain text and must be encapsulated into a html
Private Sub SetDocument(ByVal text As String, ByVal MakeFootHeader As Boolean)
   Dim ch As Long
   Dim txt As String
   If MakeFootHeader = True Then
      txt = GetHeader("")
      txt = txt & "    <PRE>" & vbCrLf
      txt = txt & text
      txt = txt & "    </PRE>" & vbCrLf
      txt = txt & GetFooter
      Else
      txt = text
      End If
   ch = FreeFile
   Open App.Path & "\temp.htm" For Output As ch
   Print #ch, txt
   Close ch
   DoNavigate App.Path & "\temp.htm"
End Sub


Private Sub Form_Load()
   App.Title = "SC-55 SysEx & Bulk dump"
   GetSC55Params App.Path & "\sc55par.dat"
   GetMacros
   SetDefaultDump
   CD_File.CancelError = True
   CD_File.InitDir = App.Path
   mMPU401OUT = 256 ' =empty
   mMPU401IN = 256
   SetChannel 0
   frmBD.Show
   
   DoNavigate App.Path & "\BulkDump.htm"
   MakePiano picKlav
   HScroll.Value = CInt(HScroll.Max * 1 / 3)

End Sub

Private Sub Form_Resize()
   If Me.WindowState = vbMinimized Then Exit Sub
   wbr.Width = Me.ScaleWidth - 2 * wbr.Left
   wbr.Height = Me.ScaleHeight - wbr.Left - wbr.Top
   If kadO.Visible = True Then
      wbr.Height = wbr.Height - kadO.Height - 60
      kadO.Top = Me.ScaleHeight - kadO.Height - 30
      kadO.Width = Me.ScaleWidth - kadO.Left * 2
      HScroll.Width = kadO.Width - 60
      HScroll.Max = picKlav.Width - kadO.Width + 90
      HScroll.LargeChange = HScroll.Max \ 10
      HScroll.Refresh
      End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   MidiTHRU_Port "close"
   MidiOUT_Port "close"
   'MidiIN_Port "close"
   End
End Sub


Private Sub HScroll_Change()
   picKlav.Left = -HScroll.Value
End Sub

Private Sub HScroll_Scroll()
   picKlav.Left = -HScroll.Value
End Sub


Private Sub mnuFileDumpSave_Click()
   CD_File.hWndOwner = Me.hwnd
   CD_File.FileName = Chr(0)
   CD_File.filter = "Midi file |*.mid|All Files (*.*)|*.*"
   CD_File.Flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
   CD_File.MaxFileSize = 255
   CD_File.DefaultExt = ".mid"
   On Error Resume Next
   CD_File.ShowSave
   If err = 32755 Then
      Exit Sub
   ElseIf err = 0 Then
      Busy Me, True
      On Error GoTo 0
      SYXdumpSave CD_File.FileName
      Busy Me, False
      If err <> 0 Then MsgBox err.Description
   Else
      MsgBox err.Description
   End If
End Sub

Private Sub mnuFileExit_Click()
   MidiTHRU_Port "close"
   MidiOUT_Port "close"
   'MidiIN_Port "close"
   End
End Sub

Private Sub mnuFileHTMLOpen_Click()
   Dim ch As Long
   CD_File.hWndOwner = Me.hwnd
   CD_File.FileName = Chr(0)
   CD_File.filter = "HTML file |*.htm;*.html|All Files (*.*)|*.*"
   CD_File.Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_EXPLORER
   CD_File.MaxFileSize = 255
   On Error Resume Next
   CD_File.ShowOpen
   If err = 32755 Then
      Exit Sub
   ElseIf err = 0 Then
      Busy Me, True
      If InStr(1, CD_File.FileName, ".htm") > 0 Then
         DoNavigate CD_File.FileName
         Else
         SetDocument File2html(CD_File.FileName, ""), False
         End If
      If err <> 0 Then MsgBox err.Description
      Busy Me, False
   Else
      MsgBox err.Description
   End If

End Sub

Private Sub mnuFilePartOpen_Click()
   Dim ch As Long
   Dim ipb As String
   Dim Part As Byte
   Dim mChan As Byte
   Dim I As Long
   
   CD_File.hWndOwner = Me.hwnd
   CD_File.FileName = Chr(0)
   CD_File.filter = "Part file |*.prt|All Files (*.*)|*.*"
   CD_File.Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_EXPLORER
   CD_File.MaxFileSize = 255
   On Error Resume Next
   CD_File.ShowOpen
   If err = 32755 Then
      Exit Sub
   ElseIf err = 0 Then
      ipb = InputBox("In which part (1-16) do you want to paste the information?", "Open Part info", "1")
      Part = Val(ipb) - 1
      If Part > 15 Then
         If MsgBox("Wrong part number? Paste the info in part 1?", vbYesNo, "Oeps") = vbNo Then Exit Sub
         Part = 0
         End If
      Busy Me, True
      ch = FreeFile
      Open CD_File.FileName For Binary As ch
      mChan = Choose(Part + 1, 1, 2, 3, 4, 5, 6, 7, 8, 9, 0, 10, 11, 12, 13, 14, 15)
      For I = 0 To 111
         Get #ch, , dmpB(72 + mChan * 112 + I)
      Next I
      Close ch
      Busy Me, False
      If err <> 0 Then MsgBox err.Description
   Else
      MsgBox err.Description
   End If


End Sub


Private Sub mnuFilePartSave_Click()
   Dim ch As Long
   Dim ipb As String
   Dim Part As Byte
   Dim mChan As Byte
   Dim I As Long

   CD_File.hWndOwner = Me.hwnd
   CD_File.FileName = Chr(0)
   CD_File.filter = "Part file |*.prt|All Files (*.*)|*.*"
   CD_File.Flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
   CD_File.MaxFileSize = 255
   CD_File.DefaultExt = ".prt"
   On Error Resume Next
   CD_File.ShowSave
   If err = 32755 Then
      Exit Sub
   ElseIf err = 0 Then
      ipb = InputBox("Which part (1-16) do you want to save?", "Save Part info", "1")
      Part = Val(ipb) - 1
      If Part > 15 Then
         If MsgBox("Wrong part number? Save the info in part 1?", vbYesNo, "Oeps") = vbNo Then Exit Sub
         Part = 0
         End If
      Busy Me, True
      ch = FreeFile
      Open CD_File.FileName For Binary As ch
      mChan = Choose(Part + 1, 1, 2, 3, 4, 5, 6, 7, 8, 9, 0, 10, 11, 12, 13, 14, 15)
      For I = 0 To 111
         Put #ch, , dmpB(72 + mChan * 112 + I)
      Next I
      Close ch
      If err <> 0 Then MsgBox err.Description
      Busy Me, False
      If err <> 0 Then MsgBox err.Description
   Else
      MsgBox err.Description
   End If

End Sub

Private Sub mnuFileMidiPlay_Click()
   Dim ch As Long
   CD_File.hWndOwner = Me.hwnd
   CD_File.FileName = Chr(0)
   CD_File.filter = "Midi file |*.mid|All Files (*.*)|*.*"
   CD_File.Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_EXPLORER
   CD_File.MaxFileSize = 255
   On Error Resume Next
   CD_File.ShowOpen
   If err = 32755 Then
      Exit Sub
   ElseIf err = 0 Then
      'On Error GoTo 0
      Busy Me, True
      SetDocument playMidiFile(CD_File.FileName), False
      Busy Me, False
      If err <> 0 Then MsgBox err.Description
   Else
      MsgBox err.Description
   End If
End Sub

Private Sub mnuFileMidiRead_Click()
   Dim ch As Long
   CD_File.hWndOwner = Me.hwnd
   CD_File.FileName = Chr(0)
   CD_File.filter = "Midi file |*.mid|All Files (*.*)|*.*"
   CD_File.Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_EXPLORER
   CD_File.MaxFileSize = 255
   On Error Resume Next
   CD_File.ShowOpen
   If err = 32755 Then
      Exit Sub
   ElseIf err = 0 Then
      'On Error GoTo 0
      frmReadMid.Show
      OK = False: Cancel = False
      While OK = False And Cancel = False: DoEvents: Wend
      If Cancel = False Then
         Busy Me, True
         SetDocument readMidiFile(CD_File.FileName), True
         Busy Me, False
         End If
      Unload frmReadMid
      If err <> 0 Then MsgBox err.Description
   Else
      MsgBox err.Description
   End If
End Sub

Private Sub mnuFileHTMLSave_Click()
   If CurHTMLfile = "" Or CurHTMLfile = "about:blank" Then MsgBox "Nothing to save yet!": Exit Sub
   CD_File.hWndOwner = Me.hwnd
   CD_File.FileName = Chr(0)
   CD_File.filter = "HTML file |*.htm;*.html|All Files (*.*)|*.*"
   CD_File.Flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
   CD_File.MaxFileSize = 255
   CD_File.DefaultExt = ".htm"
   On Error Resume Next
   CD_File.ShowSave
   If err = 32755 Then
      Exit Sub
   ElseIf err = 0 Then
      Busy Me, True
      FileCopy CurHTMLfile, CD_File.FileName
      Busy Me, False
      If err <> 0 Then MsgBox err.Description
   Else
      MsgBox err.Description
   End If
End Sub

Private Sub mnuHelpAbout_Click()
   frmAbout.Show 1, Me
End Sub

Private Sub mnuHelpInfo_Click()
   DoNavigate App.Path & "\BulkDump.htm"
End Sub


Private Sub mnuMidiDevices_Click()
   Dim OWasOpen As Boolean, IWasOpen As Boolean
   If hMidiOUT <> 0 Then OWasOpen = True: MidiOUT_Port "close"
   'If hMidiIN <> 0 Then IWasOpen = True: MidiIN_Port "close"
   frmDevCap.Show 1
   If OWasOpen = True Then MidiOUT_Port "open"
   'If IWasOpen = True Then MidiIN_Port "open"
End Sub

Private Sub mnuMidiKlavier_Click()
   If mnuMidiKlavier.Checked = True Then
      mnuMidiKlavier.Checked = False
      kadO.Visible = False
      Form_Resize
      Else
      If hMidiOUT = 0 Then MidiOUT_Port "open"
      If hMidiOUT = 0 Then Exit Sub
      mnuMidiKlavier.Checked = True
      kadO.Visible = True
      Form_Resize
      End If
End Sub


Private Sub mnuMidiPortsOff_Click()
   On Error Resume Next
   Unload frmMixLong
   Unload frmMixA
   Unload frmMixP
   On Error GoTo 0
   
   If mnuMidiThru.Checked = True Then
      mnuMidiThru.Checked = False
      MidiTHRU_Port "close"
      End If
   If mnuMidiKlavier.Checked = True Then
      mnuMidiKlavier.Checked = False
      kadO.Visible = False
      Form_Resize
      End If
   MidiOUT_Port "close"
   'MidiIN_Port "close"
End Sub

Private Sub mnuMidiSeq_Click()
   frmSeq.Show
End Sub

Private Sub mnuMidiThru_Click()
   If mnuMidiThru.Checked = False Then
      mnuMidiThru.Checked = True
      MidiTHRU_Port "open"
      Else
      mnuMidiThru.Checked = False
      MidiTHRU_Port "close"
      End If
End Sub


Private Sub mnuMixAll_Click()
   frmMixA.Show
End Sub

Private Sub mnuMixParts_Click()
   frmMixP.Show
End Sub

Private Sub mnuDumpPrint_Click()
   frmBDp.Show
End Sub

Private Sub mnuDumpAlternat_Click()
   frmAltOpt.Show 1, Me
   If OK = False Then Exit Sub
   Busy Me, True
   If SaveAlternat = False Then
      SetDocument strAlternative(UseShMsg), True
      Else
      SaveAlternative UseShMsg
      End If
   Busy Me, False
End Sub

Private Sub mnuDumpHTML_Click(Index As Integer)
   Busy Me, True
   Select Case Index
   Case 0: SetDocument strSYXDataB(0), True
   Case 1: SetDocument strSYXDataB(1), True
   End Select
   Busy Me, False
End Sub

Private Sub mnuFileDumpCompare_Click()
   CD_File.hWndOwner = Me.hwnd
   CD_File.FileName = Chr(0)
   CD_File.filter = "MIDI file |*.mid|All Files (*.*)|*.*"
   CD_File.Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_EXPLORER
   CD_File.MaxFileSize = 255
   On Error Resume Next
   CD_File.ShowOpen
   If err = 32755 Then
      Exit Sub
   ElseIf err = 0 Then
      Busy Me, True
      DumpCompare CD_File.FileName
      Busy Me, False
   Else
      MsgBox err.Description
   End If
End Sub

Private Sub mnuFileDumpOpen_Click()
   Dim ret As String
   
   CD_File.hWndOwner = Me.hwnd
   CD_File.FileName = Chr(0)
   CD_File.filter = "MIDI file |*.mid|All Files (*.*)|*.*"
   CD_File.Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_EXPLORER
   CD_File.MaxFileSize = 255
   On Error Resume Next
   CD_File.ShowOpen
   Me.Refresh
   If err = 32755 Then
      Exit Sub
   ElseIf err = 0 Then
      Busy Me, True
      If SYXdumpOpen(CD_File.FileName, ret) = True Then
         Me.Caption = "Analyse SC55 bulk-dump in " & CD_File.FileTitle
         CurDmpMidFile = CD_File.FileName
         CurDmpMidFileTitle = CD_File.FileTitle
         mnuFileDumpSave.Enabled = True
         mnuFileDumpCompare.Enabled = True
         End If
      SetDocument ret, True
      Busy Me, False
   Else
      MsgBox err.Description
   End If

End Sub


Private Sub mnuMidiSyx_Click()
   frmMixLong.Show
End Sub


Private Sub mnuMidiGSReset_Click()
   GSResetAll
End Sub

Private Sub picKlav_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim Oct As Long
   Dim No As Long
   Dim mX As Single
   
   Oct = X \ (16 * 7)
   If picKlav.Point(X, Y) = 0 And Y < 41 Then
      mX = X - 10
      No = Oct * 12 + Choose(((mX \ 16) Mod 7) + 1, 1, 3, 5, 6, 8, 10, 11)
      Else
      mX = X
      No = Oct * 12 + Choose(((mX \ 16) Mod 7) + 1, 0, 2, 4, 5, 7, 9, 11)
      End If
   
   midiMessageOut = NOTE_ON + CurChannel
   midiData1 = No
   midiData2 = 120
   SendMidiShortOut
   CurKeyID = No
   lblKey.Left = X - 14
   lblKey.Caption = isNote(No)
   lblKey.Visible = True
End Sub

Private Sub picKlav_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   midiMessageOut = NOTE_OFF + CurChannel
   midiData1 = CurKeyID
   midiData2 = 120
   SendMidiShortOut
   lblKey.Visible = False
End Sub

Private Sub wbr_DocumentComplete(ByVal pDisp As Object, URL As Variant)
   DocComplete = True
   CurHTMLfile = URL
End Sub


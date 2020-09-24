VERSION 5.00
Begin VB.Form frmBDp 
   Caption         =   "Print SC55 - bulk dump"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "frmBDp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   2100
      TabIndex        =   10
      Top             =   0
      Width           =   1815
      Begin VB.OptionButton optVal 
         Caption         =   "Norm"
         Height          =   315
         Index           =   1
         Left            =   780
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   60
         Width           =   615
      End
      Begin VB.OptionButton optVal 
         Caption         =   "Hex"
         Height          =   315
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.OptionButton optPag 
      Caption         =   "Pag 2"
      Height          =   315
      Index           =   1
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   60
      Width           =   975
   End
   Begin VB.OptionButton optPag 
      Caption         =   "Pag 1"
      Height          =   315
      Index           =   0
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   5820
      TabIndex        =   6
      Text            =   "SC55 - bulk dump"
      Top             =   60
      Width           =   3555
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4020
      TabIndex        =   5
      Top             =   60
      Width           =   1155
   End
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
      Height          =   2790
      Left            =   60
      ScaleHeight     =   2730
      ScaleWidth      =   2295
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   420
      Width           =   2355
      Begin VB.PictureBox kadI 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1590
         Left            =   0
         ScaleHeight     =   1590
         ScaleWidth      =   1875
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   1875
         Begin VB.PictureBox picBMP 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1200
            Left            =   45
            ScaleHeight     =   80
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   97
            TabIndex        =   4
            Top             =   45
            Width           =   1455
         End
      End
      Begin VB.HScrollBar HScroll 
         Height          =   270
         Left            =   0
         Min             =   -45
         TabIndex        =   2
         Top             =   2460
         Width           =   1050
      End
      Begin VB.VScrollBar VScroll 
         Height          =   915
         Left            =   1995
         Min             =   -45
         TabIndex        =   1
         Top             =   15
         Width           =   270
      End
   End
   Begin VB.Label lbl 
      Caption         =   "Title"
      Height          =   255
      Left            =   5340
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmBDp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form is a kind of a preview to print the dump data.
Option Explicit
Dim CurPage As Integer
Dim CurValStyle As Integer ' hex/dec

Dim tW As Long, tH As Long ' text width/height
Dim X1 As Long, X2 As Long
Dim Y1 As Long, Y2 As Long

' general routine for pictures with scrollbars
' container should be in twips mode
Private Sub CheckScrolls(pic As PictureBox, _
                        kadI As PictureBox, kadO As PictureBox, _
                        HS As HScrollBar, VS As VScrollBar)
   kadI.Width = kadO.Width - 60
   VS.Left = kadI.Width - VS.Width + 15
   HS.Width = kadI.Width - VS.Width
   
   kadI.Height = kadO.Height - 60
   HS.Top = kadI.Height - HS.Height + 15
   VS.Height = kadI.Height - HS.Height
   
   If pic.Width > kadI.Width - 90 Then
      kadI.Height = kadO.Height - HS.Height - 90: VS.Height = kadI.Height
      HS.Max = pic.Width - kadI.Width + 90: HS.Visible = True
      Else
      VS.Height = kadI.Height
      HS.Value = HS.Min: HS.Visible = False
      End If
   If pic.Height > kadI.Height - 90 Then
      kadI.Width = kadO.Width - VS.Width - 90: HS.Width = kadI.Width
      VS.Max = pic.Height - kadI.Height + 90: VS.Visible = True
      Else
      HS.Width = kadI.Width
      VS.Value = VS.Min: VS.Visible = False
      End If
   If pic.Width > kadI.Width - 90 And pic.Height > kadI.Height - 90 Then
      kadI.Height = kadO.Height - HS.Height - 90: VS.Height = kadI.Height
      HS.Max = pic.Width - kadI.Width + 90: HS.Visible = True
      kadI.Width = kadO.Width - VS.Width - 90: HS.Width = kadI.Width
      VS.Max = pic.Height - kadI.Height + 90: VS.Visible = True
      End If
   VS.Refresh
   HS.Refresh
End Sub

' print text, marked or not, hex/dec
Private Sub PrintAt(pic As Control, X As Long, Y As Long, _
            ByVal txt As String, Marked As Boolean, _
            ParType As Byte)
   pic.CurrentX = X
   pic.CurrentY = Y
   If Marked Then
      pic.FontBold = True
      pic.FontItalic = True
      pic.ForeColor = QBColor(9)
      Else
      pic.FontBold = False
      pic.FontItalic = False
      pic.ForeColor = 0
      End If
   If CurValStyle = 1 Then txt = isValue(txt, ParType)
   pic.Print txt
End Sub

Private Sub ShowRxSwitches(pic As Control)
   Dim I As Long, J As Long
   Dim s(16, 16) As Byte
   Dim B As Byte
   Dim txt As String
   Dim Marked As Boolean
   
   If CurValStyle = 0 Then
      For I = 36 To 37
         PrintAt pic, X1, Y1, PPar(I).ShortName, False, 0
         For J = 0 To 15
            B = dmpB(72 + J * 112 + PPar(I).ByteOffs)
            txt = HexByte(B)
            Marked = IIf(B = PPar(I).Default, False, True)
            PrintAt pic, X2 + J * tW, Y1, txt, Marked, PPar(I).Type
         Next J
         Y1 = Y1 + tH
      Next I
      Exit Sub
      End If
   For J = 0 To 15
      B = dmpB(72 + J * 112 + PPar(36).ByteOffs)
      B = RotateLByte(B)
      For I = 0 To 7
         s(I, J) = IIf((B And (2 ^ (7 - I))) = 0, 0, 1)
      Next I
      B = dmpB(72 + J * 112 + PPar(37).ByteOffs)
      B = RotateRByte(B)
      For I = 0 To 7
         s(I + 8, J) = IIf((B And (2 ^ (I))) = 0, 0, 1)
      Next I
   Next J
   For I = 0 To 15
      B = s(I, 7): s(I, 7) = s(I, 15): s(I, 15) = B
   Next I
   For I = 0 To 15
      txt = "Rx." & Choose(I + 1, "Bend", "Caf", "PrgCh", "CtlCh", "Paf", "Note", "RPN", "NRPN", "Modul", "Volum", "Pan", "Expr", "Hold", "Port", "Sost", "Soft")
      PrintAt pic, X1, Y1, txt, False, 0
      For J = 0 To 15
         txt = IIf(s(I, J) = 0, "Off", "On")
         PrintAt pic, X2 + J * tW, Y1, txt, IIf(s(I, J) = 1, False, True), 0
      Next J
      Y1 = Y1 + tH
   Next I
End Sub

Private Sub ShowSYX(pic As Control, ByVal Page As Integer)
   Dim iB As Byte
   Dim I As Long, J As Long
   Dim kol1 As Long, kol2 As Long
   Dim txt As String, Marked As Boolean
   Dim W As Long, H As Long
   Dim ch As Long
   
   W = pic.ScaleWidth
   H = pic.ScaleHeight
   pic.FontName = "Arial": pic.FontSize = 8
   tW = pic.TextWidth("0000000")
   tH = pic.TextHeight("W")
   X1 = W * 0.04: X2 = W * 0.25: kol1 = W / 2: kol2 = W * 0.4
   Y1 = H * 0.02
   
   txt = txtTitle.text & " - pag. " & CStr(Page + 1)
   PrintAt pic, (pic.ScaleWidth - pic.TextWidth(txt)) / 2, Y1, txt, False, 0
   txt = ""
   Y1 = Y1 + tH
   
   ' common
   If Page = 0 Then
      I = 4
      PrintAt pic, X1, Y1, "Patch name :", False, 0
      For J = APar(I).ByteOffs To APar(I).ByteOffs + 15
         txt = txt & Chr(dmpB(J))
      Next J
      PrintAt pic, X2, Y1, txt, False, 0
      
      Y1 = Y1 + tH * 2: I = 0
      PrintAt pic, X1, Y1, "Master tune :", False, 0
      txt = HexByte(dmpB(APar(I).ByteOffs)) & " " & HexByte(dmpB(APar(I).ByteOffs + 1))
      Marked = IIf(dmpB(APar(I).ByteOffs) = 4 And dmpB(APar(I).ByteOffs + 1) = 0, False, True)
      PrintAt pic, X2, Y1, txt, Marked, APar(I).Type
   
      I = 2
      PrintAt pic, X1 + kol1, Y1, "Master key shift :", False, 0
      txt = HexByte(dmpB(APar(I).ByteOffs))
      Marked = IIf(dmpB(APar(I).ByteOffs) = APar(I).Default, False, True)
      PrintAt pic, X2 + kol1, Y1, txt, Marked, APar(I).Type
      
      Y1 = Y1 + tH: I = 1
      PrintAt pic, X1, Y1, "Master volume :", False, 0
      txt = HexByte(dmpB(APar(I).ByteOffs))
      Marked = IIf(dmpB(APar(I).ByteOffs) = APar(I).Default, False, True)
      PrintAt pic, X2, Y1, txt, Marked, APar(I).Type
        
      I = 3
      PrintAt pic, X1 + kol1, Y1, "Master panpot :", False, 0
      txt = HexByte(dmpB(APar(I).ByteOffs))
      Marked = IIf(dmpB(APar(I).ByteOffs) = APar(I).Default, False, True)
      PrintAt pic, X2 + kol1, Y1, txt, Marked, APar(I).Type
   
      Y1 = Y1 + tH * 2
      For I = 13 To AParCount - 1
         PrintAt pic, X1, Y1, Trim(APar(I).name) & " :", False, 0
         txt = HexByte(dmpB(APar(I).ByteOffs))
         Marked = IIf(dmpB(APar(I).ByteOffs) = APar(I).Default, False, True)
         PrintAt pic, X2, Y1, txt, Marked, APar(I).Type
         If I < AParCount - 1 Then
            PrintAt pic, X1 + kol1, Y1, Trim(APar(I - 7).name) & " :", False, 0
            txt = HexByte(dmpB(APar(I - 7).ByteOffs))
            Marked = IIf(dmpB(APar(I - 7).ByteOffs) = APar(I - 7).Default, False, True)
            PrintAt pic, X2 + kol1, Y1, txt, Marked, APar(I - 7).Type
            End If
         Y1 = Y1 + tH
      Next I
      ' patches
      Y1 = Y1 + tH: I = 1
      For J = 0 To 3
         ch = isChannel(J) + 1
         txt = Format(ch, "00") & " " & getPatchName(J)
         PrintAt pic, X1, Y1, txt, False, 0
         ch = isChannel(J + 4) + 1
         txt = Format(ch, "00") & " " & getPatchName(J + 4)
         PrintAt pic, X1 + kol2 / 2, Y1, txt, False, 0
         ch = isChannel(J + 8) + 1
         txt = Format(ch, "00") & " " & getPatchName(J + 8)
         PrintAt pic, X1 + kol2, Y1, txt, False, 0
         ch = isChannel(J + 12) + 1
         txt = Format(ch, "00") & " " & getPatchName(J + 12)
         PrintAt pic, X1 + kol2 * 3 / 2, Y1, txt, False, 0
         Y1 = Y1 + tH
      Next J
      ' parts partial reserve
      X2 = X1 + tW * 1.4
      Y1 = Y1 + tH: I = 5
      PrintAt pic, X1, Y1, APar(I).ShortName, False, 0
      For J = 0 To 15
         iB = dmpB(J + APar(I).ByteOffs)
         txt = HexByte(iB)
         Marked = IIf(iB = Val(Mid("2622222222000000", J + 1, 1)), False, True)
         PrintAt pic, X2 + J * tW, Y1, txt, Marked, APar(I).Type
      Next J
      Y1 = Y1 + tH
      End If

   ' parts koppen
   X2 = X1 + tW * 1.4
   Y1 = Y1 + tH
   PrintAt pic, X1, Y1, "PART", True, 0
   For J = 0 To 15
      txt = Choose(J + 1, "10", "1", "2", "3", "4", "5", "6", "7", "8", "9", "11", "12", "13", "14", "15", "16")
      PrintAt pic, X2 + J * tW, Y1, txt, True, 0
   Next J
   ' parts parameters
   Y1 = Y1 + tH
   If Page = 0 Then
   
      For I = 0 To PParCount - 1
         If PPar(I).ShortName = "RX1 0/1" Then Exit For
         PrintAt pic, X1, Y1, PPar(I).ShortName, False, 0
         For J = 0 To 15
            iB = dmpB(72 + J * 112 + PPar(I).ByteOffs)
            txt = HexByte(iB)
            If PPar(I).ShortName = "CHANNEL" Then
               Marked = IIf(iB = Val("&H" & Mid("9012345678ABCDEF", J + 1, 1)), False, True)
               Else
               Marked = IIf(iB = PPar(I).Default, False, True)
               End If
            If J = 0 And I = 3 And iB = &HB0 Then Marked = False
            PrintAt pic, X2 + J * tW, Y1, txt, Marked, PPar(I).Type
         Next J
         Y1 = Y1 + tH
      Next I
      ShowRxSwitches pic
      
      Else
      
      J = 0: While PPar(J).ShortName <> "M PITCH": J = J + 1: Wend
      For I = J To PParCount - 1
         PrintAt pic, X1, Y1, PPar(I).ShortName, False, 0
         For J = 0 To 15
            iB = dmpB(72 + J * 112 + PPar(I).ByteOffs)
            txt = HexByte(iB)
            Marked = IIf(iB = PPar(I).Default, False, True)
            PrintAt pic, X2 + J * tW, Y1, txt, Marked, PPar(I).Type
         Next J
         Y1 = Y1 + tH
      Next I
      
      End If
End Sub

Private Sub cmdPrint_Click()
   Dim msg As String
   Dim ret As Variant
   
   msg = "Choose" & vbCrLf
   msg = msg & "Yes to print both pages" & vbCrLf
   msg = msg & "No to print only the current page" & CStr(CurPage + 1) & vbCrLf
   ret = MsgBox(msg, vbYesNoCancel, "Print")
   If ret = vbCancel Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   Printer.ScaleMode = picBMP.ScaleMode
   Printer.Font = picBMP.Font
   If ret = vbYes Then
      ShowSYX Printer, 0
      Printer.NewPage
      ShowSYX Printer, 1
      Else
      ShowSYX Printer, CurPage
      End If
   Printer.EndDoc
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   picBMP.Width = Printer.Width
   picBMP.Height = Printer.Height
   txtTitle.text = txtTitle.text & " - " & CurDmpMidFileTitle
   Me.Show
   DoEvents
   CurValStyle = 0 ' hex
   optPag(0).Value = True ' 1st page
End Sub

Private Sub Form_Resize()
   If Me.WindowState = vbMinimized Then Exit Sub
   kadO.Width = Me.ScaleWidth - kadO.Left * 2
   kadO.Height = Me.ScaleHeight - kadO.Top - kadO.Left
   CheckScrolls picBMP, kadI, kadO, HScroll, VScroll
End Sub

Private Sub HScroll_Change()
   picBMP.Left = -HScroll.Value
End Sub

Private Sub HScroll_Scroll()
   picBMP.Left = -HScroll.Value
End Sub

Private Sub optPag_Click(Index As Integer)
   CurPage = Index
   Screen.MousePointer = vbHourglass
   picBMP.Cls
   ShowSYX picBMP, CurPage
   Screen.MousePointer = vbDefault
End Sub

Private Sub optVal_Click(Index As Integer)
   CurValStyle = Index
   Screen.MousePointer = vbHourglass
   picBMP.Cls
   ShowSYX picBMP, CurPage
   Screen.MousePointer = vbDefault
End Sub

Private Sub VScroll_Change()
   picBMP.Top = -VScroll.Value
End Sub

Private Sub VScroll_Scroll()
   picBMP.Top = -VScroll.Value
End Sub


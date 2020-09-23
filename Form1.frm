VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5535
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VirtualTypewriter.ucCarriageGuard ucCarriageGuard1 
      Height          =   735
      Left            =   2100
      TabIndex        =   2
      Top             =   4800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
   End
   Begin VB.PictureBox pctPaper 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Aeolus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3075
      Left            =   1080
      ScaleHeight     =   3075
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   1080
      Width           =   3555
   End
   Begin VB.PictureBox pctRoller 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   480
      ScaleHeight     =   1440
      ScaleWidth      =   5205
      TabIndex        =   1
      Top             =   4080
      Width           =   5205
   End
   Begin VB.Image imgTypewriter 
      Appearance      =   0  'Flat
      Height          =   2085
      Left            =   240
      Picture         =   "Form1.frx":0000
      Top             =   3480
      Width           =   5700
   End
   Begin VB.Image imgRoller 
      Height          =   3450
      Left            =   1800
      Picture         =   "Form1.frx":087D
      Top             =   300
      Visible         =   0   'False
      Width           =   3675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Virtual Typewriter
'  Characters a printed within a PictureBox.

Option Explicit

Const defaultSpacing = "X"     ' default size of
Dim defaultCharacterWidth As Integer
Dim defaultCharacterHeight As Integer

Dim displayCenterX As Integer   ' \
Dim displayCenterY As Integer   ' / center of impact area (X,Y)

Dim pageLineMax As Integer     ' max lines on page
Dim pageLineCount As Integer   ' line number of page
Dim lineCharMax As Integer     ' Maximum characters on a line
Dim lineCharCount As Integer   ' character number of line
Dim lastLine As Boolean        ' marks when lsat line is reached

Dim marginBell As Boolean      ' reached edge of paper

Dim marginTop As Integer       ' top margin (in lines)
Dim marginLeft As Integer      ' left margin (in chars)
Dim marginRight As Integer     ' right margin (in chars)
Dim marginBottom As Integer    ' bottom margin (in lines)

Private Sub PositionPaper()
 With pctPaper
  .CurrentX = defaultCharacterWidth * lineCharCount  ' \
  .CurrentY = defaultCharacterHeight * pageLineCount ' / reset Cursor Position (X,Y)
  .Move displayCenterX - .CurrentX, displayCenterY - .CurrentY
 End With
 
 With Printer
  .CurrentX = defaultCharacterWidth * lineCharCount  ' \
  .CurrentY = defaultCharacterHeight * pageLineCount ' / reset Cursor Position (X,Y)
 End With
 
 With pctRoller
  .CurrentX = defaultCharacterWidth * (lineCharCount + 2)
  .Move displayCenterX - .CurrentX, Me.ScaleHeight - .Height
 End With
 
 MakeTransparent Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim tx As Integer
 Dim moved As Boolean
 
 moved = True
 With pctPaper
  Select Case KeyCode
   
   Case vbKeyUp
    PlaySound App.Path & "\roller.wav", 0, 1
    pageLineCount = Greater(pageLineCount - 1, 0)
   ' End Case vbKeyUp
   
   Case vbKeyDown
    PlaySound App.Path & "\roller.wav", 0, 1
    pageLineCount = Lesser(pageLineCount + 1, pageLineMax)
   ' End Case vbKeyDown
   
   Case vbKeyLeft
    PlaySound App.Path & "\back space.wav", 0, 1
    lineCharCount = Greater(lineCharCount - 1, 0)
   ' End Case vbKeyDown
   
   Case vbKeyRight
    PlaySound App.Path & "\space bar.wav", 0, 1
    lineCharCount = Lesser(lineCharCount + 1, lineCharMax)
   ' End Case vbKeyDown
   
   Case vbKeyPageUp
    PlaySound App.Path & "\carriage return.wav", 0, 1
    pageLineCount = 0
   ' End Case vbKeyPageDown
   
   Case vbKeyPageDown
    PlaySound App.Path & "\carriage return.wav", 0, 1
    pageLineCount = pageLineMax
   ' End Case vbKeyPageDown
   
   Case Else
    moved = False
   ' End Case Else
   
  End Select
 End With
 
 If moved Then
  Call PositionPaper
 End If
  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 Dim tx As Integer
 Dim ch As String
 Dim finished As Boolean
 
 finished = False
 
 With pctPaper
  Select Case KeyAscii
   Case vbKeyEscape
    finished = True
   ' End Case vbKeyEscape
   
   Case vbKeyBack
    PlaySound App.Path & "\back space.wav", 0, 1
    lineCharCount = Greater(lineCharCount - 1, 0)
   ' End Case vbKeyBack
   
   Case vbKeyReturn
    PlaySound App.Path & "\carriage return.wav", 0, 1
    tx = lineCharCount
    pageLineCount = Lesser(pageLineCount + 1, pageLineMax)
    For lineCharCount = tx To 1 Step -1
     Call PositionPaper
    Next lineCharCount
   ' End Case vbKeyReturn
   
   Case Else               'everything else
    ch = Chr(KeyAscii)
    If ch = " " Then
     PlaySound App.Path & "\space bar.wav", 0, 1
    Else
     PlaySound App.Path & "\key hit.wav", 0, 1
    End If
    pctPaper.Print ch;                 ' this alters .currentX/Y
    lineCharCount = Lesser(lineCharCount + 1, lineCharMax)
  
    ' check for right margin and bell is necessary
    If lineCharCount > marginRight Then
     If Not marginBell Then
      marginBell = True
      PlaySound App.Path & "\margin bell.wav", 0, 0
     End If
    Else
     marginBell = False
    End If
   ' End Case Else
  End Select
 End With
 Call PositionPaper
  
 If finished Then Unload Me
End Sub

Private Sub Form_Load()
 
' Printer.Print ""                        ' begin printer stuff
 With Printer
  .FontName = "Courier New"
  .FontSize = 14
  Me.ScaleMode = .ScaleMode        ' make both the same
  pctPaper.Width = .ScaleWidth
  pctPaper.Height = .ScaleHeight
  pctPaper.FontName = .FontName
  pctPaper.FontSize = .FontSize
 End With
 
 Me.ScaleMode = 1                     ' set to TWIPS for placement
 With pctPaper
  defaultCharacterWidth = .TextWidth(defaultSpacing)
  defaultCharacterHeight = .TextHeight(defaultSpacing)
  
  pageLineMax = Int(.ScaleHeight / defaultCharacterHeight) - 1
  lineCharMax = Int(.ScaleWidth / defaultCharacterWidth) - 1
 End With
 
 With pctRoller
  .Width = pctPaper.Width + 4 * defaultCharacterWidth
  .PaintPicture imgRoller.Picture, 0, 0, .ScaleWidth, .ScaleHeight
 End With
 
 lastLine = False
 pageLineCount = 0
 lineCharCount = 0
 marginBell = False                        ' reset bell
 marginLeft = 0                            ' margin to min
 marginRight = lineCharMax - 7             ' margin to max
 marginTop = 0
 marginBottom = pageLineMax
 
 Me.BorderStyle = 0
 Me.WindowState = 2
 Me.Show

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Dim response As Integer
 
 response = MsgBox("Print Page From Virtual Typewriter?", vbYesNoCancel, "Ending Virtual Typewriter")
  
 If response = vbCancel Then       ' cancel ending
  Cancel = True
 Else
  If response = vbYes Then        ' print before ending?
   Printer.Print ""
   Printer.PaintPicture pctPaper.Image, 0, 0
   Printer.NewPage
   Printer.EndDoc
  End If
  
  Cancel = False                  ' end program
 End If
 
End Sub

Private Sub Form_Resize()
 If Me.WindowState = vbMinimized Then Exit Sub
 
 displayCenterX = Me.ScaleWidth / 2
 With ucCarriageGuard1
  .Move displayCenterX - .Width / 2, Me.ScaleHeight - .Height
  displayCenterY = .Top + defaultCharacterHeight / 2
  displayCenterX = displayCenterX - defaultCharacterWidth / 2
 End With
 
 With imgTypewriter
  .Move displayCenterX - .Width / 2, Me.ScaleHeight - .Height
 End With
 
 Call PositionPaper
 
End Sub


VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transparent Menu - ©2002 Wilksey!"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "Transparent Menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Transparent Menu.frx":0442
   ScaleHeight     =   3195
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox imgMenu 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "Transparent Menu.frx":C9DC
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   20
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Timer tmrMenuItem 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   2400
   End
   Begin VB.Timer tmrMenu 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2640
      Top             =   2400
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtOpacity 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   16
      Text            =   "255"
      Top             =   1200
      Width           =   375
   End
   Begin VB.HScrollBar hscrOpacity 
      Height          =   255
      Left            =   2880
      Max             =   255
      TabIndex        =   14
      Top             =   1560
      Width           =   1695
   End
   Begin VB.PictureBox picMnuFile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      ScaleHeight     =   119
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   55
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   855
      Begin VB.Label mnuItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Close"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   12
         Top             =   1560
         Width           =   495
      End
      Begin VB.Image imgItem 
         Height          =   240
         Index           =   4
         Left            =   0
         Picture         =   "Transparent Menu.frx":10C1B
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label mnuItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Find..."
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   11
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label mnuItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Open"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   120
         Width           =   855
      End
      Begin VB.Image imgItem 
         Height          =   240
         Index           =   0
         Left            =   0
         Picture         =   "Transparent Menu.frx":10D1D
         Top             =   120
         Width           =   240
      End
      Begin VB.Label mnuItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Save"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   495
      End
      Begin VB.Image imgItem 
         Height          =   240
         Index           =   1
         Left            =   0
         Picture         =   "Transparent Menu.frx":10E1F
         Top             =   480
         Width           =   240
      End
      Begin VB.Label mnuItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Print"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   495
      End
      Begin VB.Image imgItem 
         Height          =   240
         Index           =   2
         Left            =   0
         Picture         =   "Transparent Menu.frx":10F21
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgItem 
         Enabled         =   0   'False
         Height          =   240
         Index           =   3
         Left            =   0
         Picture         =   "Transparent Menu.frx":11023
         Top             =   1200
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdPicture2 
      Caption         =   "Picture"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdColor2 
      Caption         =   "Color"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Color"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdPicture 
      Caption         =   "Picture"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   2760
      Width           =   735
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   313
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File"
         Height          =   195
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Label lblImage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Picture Is This:"
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   2400
      Width           =   1380
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Set the picture or Color property below to see the effect of transparent menus."
      Height          =   435
      Left            =   1080
      TabIndex        =   18
      Top             =   600
      Width           =   3360
   End
   Begin VB.Label lblOpacity 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opacity Value:"
      Height          =   195
      Left            =   3000
      TabIndex        =   15
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label lblSubItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu 1"
      Height          =   195
      Left            =   4080
      TabIndex        =   6
      Top             =   1920
      Width           =   540
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Main"
      Height          =   195
      Left            =   3360
      TabIndex        =   5
      Top             =   1920
      Width           =   345
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Transparent Menu - ©2002 Wilksey!
'
Option Explicit
'Declarations

Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Const AC_SRC_OVER = &H0
Const SRCCOPY = &HCC0020

Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Dim mnuFileOpened As Boolean
Dim I, A As Long
Dim Opacity As Integer
Dim BF As BLENDFUNCTION, lBF As Long
Dim PictureMain As Boolean

Private Sub cmdColor_Click()
If mnuFileOpened Then
    'enabled timer
    tmrMenu.Enabled = True
    'set varaible
    mnuFileOpened = False
End If
'change the backcolor of the control
    picMenu.BackColor = vbRed
    'clear the pic box
    picMenu.Cls
If PictureMain Then
    'change the Picture of the control
    StretchBlt picMenu.hdc, 0, 0, picMenu.ScaleWidth, picMenu.ScaleHeight, imgMenu.hdc, 0, 0, imgMenu.ScaleWidth, imgMenu.ScaleHeight, SRCCOPY
End If

With BF
    .BlendOp = AC_SRC_OVER
    .BlendFlags = 0
    .SourceConstantAlpha = Opacity
    .AlphaFormat = 0
End With
'api calls
    RtlMoveMemory lBF, BF, 4
    AlphaBlend picMenu.hdc, 0, 0, picMenu.ScaleWidth, picMenu.ScaleHeight, Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, lBF
'refresh pic box
    picMenu.Refresh
End Sub

Private Sub cmdColor2_Click()
If mnuFileOpened Then
'enabled timer
    tmrMenu.Enabled = True
'set varaible
    mnuFileOpened = False
End If
'change the backcolor of the control
    picMnuFile.BackColor = vbRed
'clear the pic box
    picMnuFile.Cls
With BF
    .BlendOp = AC_SRC_OVER
    .BlendFlags = 0
    .SourceConstantAlpha = Opacity
    .AlphaFormat = 0
End With
'api calls
    RtlMoveMemory lBF, BF, 4
    AlphaBlend picMnuFile.hdc, 0, 0, picMnuFile.ScaleWidth, picMnuFile.ScaleHeight, Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, lBF
'refresh pic box
    picMnuFile.Refresh
End Sub

Private Sub cmdPicture_Click()
If mnuFileOpened Then
    'enabled timer
    tmrMenu.Enabled = True
    'set varaible
    mnuFileOpened = False
End If
'Sets variable
    PictureMain = True
'clear the pic box
    picMenu.Cls
'change the Picture of the control
    StretchBlt picMenu.hdc, 0, 0, picMenu.ScaleWidth, picMenu.ScaleHeight, imgMenu.hdc, 0, 0, imgMenu.ScaleWidth, imgMenu.ScaleHeight, SRCCOPY
With BF
    .BlendOp = AC_SRC_OVER
    .BlendFlags = 0
    .SourceConstantAlpha = Opacity
    .AlphaFormat = 0
End With
'api calls
    RtlMoveMemory lBF, BF, 4
    AlphaBlend picMenu.hdc, 0, 0, picMenu.ScaleWidth, picMenu.ScaleHeight, Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, lBF
'refresh pic box
    picMenu.Refresh
End Sub

Private Sub cmdPicture2_Click()
If mnuFileOpened Then
'enabled timer
    tmrMenu.Enabled = True
'set varaible
    mnuFileOpened = False
End If
'change the Picture of the control
    picMnuFile.Picture = imgMenu.Picture
'clear the pic box
    picMnuFile.Cls
With BF
    .BlendOp = AC_SRC_OVER
    .BlendFlags = 0
    .SourceConstantAlpha = Opacity
    .AlphaFormat = 0
End With
'api calls
    RtlMoveMemory lBF, BF, 4
    AlphaBlend picMnuFile.hdc, 0, 0, picMnuFile.ScaleWidth, picMnuFile.ScaleHeight, Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, lBF
'refresh pic box
    picMnuFile.Refresh
End Sub

Private Sub cmdReset_Click()
'Hides picMnuFile in case it is visible
    picMnuFile.Visible = False
    mnuFileOpened = False
'Sets the picture in picMenu to nothing/no picture
    picMenu.Picture = Nothing
'sets the picture box color back to form background
    picMenu.BackColor = Me.BackColor
'Sets the picture in picMnuFile to nothing/nopicture
    picMnuFile.Picture = Nothing
'sets the picture box color back to form background
    picMnuFile.BackColor = Me.BackColor
'Sets variable
    PictureMain = False
End Sub

Private Sub Form_Click()
'if condition
If mnuFileOpened Then
    'set label text to label text - 1st char
    lblFile.Caption = Mid$(lblFile.Caption, 2, Len(lblFile.Caption) - 1)
    'enable Timer
    tmrMenu.Enabled = True
    'set variable
    mnuFileOpened = False
'end of if statement
End If
End Sub

Private Sub Form_Load()
'Set Picturebox width to Form width
    picMenu.Width = Me.Width
'Set text box value
    txtOpacity = 0
'for loop
For I = 0 To mnuItem.UBound 'Ubound is the highest array value, like if u had 3 labels, lbl(0), lbl(1), lbl(2) ubound would be 2 and lbound would be 0
    'Sets mnuItem tag to mnuItem Top
    mnuItem(I).Tag = mnuItem(I).Top
'continues for loop
Next I
'Sets forms autoredraw property to true
    Me.AutoRedraw = True
'Sets picture box autoredraw property to true
    picMenu.AutoRedraw = True
'Sets the form scale mode to pixels
    Me.ScaleMode = vbPixels
'Sets the pic box scale mode to pixels
    picMenu.ScaleMode = vbPixels
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Sets label forecolor back to black
    lblFile.ForeColor = vbBlack
'for loop
For A = 0 To mnuItem.Count - 1
'Sets label forecolor back to black
    mnuItem(A).ForeColor = vbBlack
'continues for loop
Next A
End Sub

Private Sub Form_Unload(Cancel As Integer)
'End the program
    End
End Sub

Private Sub hscrOpacity_Change()
'call sub
    Call Form_Click
'Sets opacity text box to scroll bar value
    txtOpacity = hscrOpacity.Value
'sets opacity variable
    Opacity = hscrOpacity.Value
'Clear picMenu
    picMenu.Cls
If PictureMain Then
'change the Picture of the control
    StretchBlt picMenu.hdc, 0, 0, picMenu.ScaleWidth, picMenu.ScaleHeight, imgMenu.hdc, 0, 0, imgMenu.ScaleWidth, imgMenu.ScaleHeight, SRCCOPY
End If

'hooks the program to write bf's propertys
With BF
    .BlendOp = AC_SRC_OVER
    .BlendFlags = 0
    .SourceConstantAlpha = Opacity
    .AlphaFormat = 0
'end control hook
End With
    'API Calls
    RtlMoveMemory lBF, BF, 4
    AlphaBlend picMenu.hdc, 0, 0, picMenu.ScaleWidth, picMenu.ScaleHeight, Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, lBF
'Refresh the picture box
    picMenu.Refresh
End Sub

Private Sub hscrOpacity_Scroll()
'call sub
    Call hscrOpacity_Change
End Sub

Private Sub imgItem_Click(Index As Integer)
'if condition
If mnuItem(Index).Enabled = True Then
    'Call another sub
    Call mnuItem_Click(Index)
'end of if condition
End If
End Sub

Private Sub imgItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'call another sub
    Call mnuItem_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub imgItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'call another sub
    Call mnuItem_MouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub imgItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'call another sub
    Call mnuItem_MouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub imgMenu_Click()
'call sub
    Call Form_Click
End Sub

Private Sub imgMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'call sub
    Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblFile_Click()
If Not mnuFileOpened Then
'Sets the text of the label + "&"(underline letter)
    lblFile.Caption = "&" + lblFile.Caption
'makes the picture box visible
    picMnuFile.Visible = True
'Makes the form automatically redraw itself when it has been changed
    Me.AutoRedraw = True
'Makes the picture box automatically redraw itself when it has been changed
    picMnuFile.AutoRedraw = True
'Sets the scalemode of the picture box to pixels
    Me.ScaleMode = vbPixels
'Sets the scalemode of the picture box to pixels
    picMnuFile.ScaleMode = vbPixels
'hooks the program to write bf's propertys
With BF
    .BlendOp = AC_SRC_OVER
    .BlendFlags = 0
    .SourceConstantAlpha = 255 'Opacity
    .AlphaFormat = 0
'end control hook
End With
    'API Calls
    RtlMoveMemory lBF, BF, 4
    AlphaBlend picMnuFile.hdc, 0, 0, picMnuFile.ScaleWidth, picMnuFile.ScaleHeight, Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, lBF
'Enable timer
    tmrMenuItem.Enabled = True
'Set variable to true
    mnuFileOpened = True
End If
End Sub

Private Sub lblFile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Sets label forecolor to Blue
    lblFile.ForeColor = vbBlue
'If Condition Statement
If lblFile.Top < 9 And Not mnuFileOpened Then
    'set label top +8
    lblFile.Top = lblFile.Top + 3
'End of condition
End If
End Sub

Private Sub lblFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Sets label forecolor to blue
    lblFile.ForeColor = vbBlue
'for loop
For A = 0 To mnuItem.Count - 1
    'Sets label forecolor back to black
    mnuItem(A).ForeColor = vbBlack
'continue for loop
Next A
End Sub

Private Sub lblFile_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Sets label forecolor to Blue
    lblFile.ForeColor = vbBlue
'If Condition Statement
If lblFile.Top < 9 And Not mnuFileOpened Then
    'set label top -8
    lblFile.Top = lblFile.Top - 3
'End of condition
End If
End Sub

Private Sub lblImage_Click()
'call sub
    Call Form_Click
End Sub

Private Sub lblImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Call sub
    Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblMain_Click()
'call sub
    Call Form_Click
End Sub

Private Sub lblMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'call sub
    Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblNote_Click()
'Call Sub
    Call Form_Click
End Sub

Private Sub lblNote_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Call Sub
    Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblOpacity_Click()
'call sub
    Call Form_Click
End Sub

Private Sub lblOpacity_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'call sub
    Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblSubItem_Click()
'call sub
    Call Form_Click
End Sub

Private Sub lblSubItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'call sub
    Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub mnuItem_Click(Index As Integer)
'Declaration
Dim Ret As Long
'Hide menu
    picMnuFile.Visible = False
    mnuFileOpened = False
'set the caption of the file menu
    lblFile.Caption = Mid$(lblFile.Caption, 2, Len(lblFile.Caption) - 1)
'Evaluate value of Index
Select Case Index
'index = 0
Case 0
    'Display a message box
    MsgBox "Open", vbOKOnly + vbInformation, "You Chose..."
'index is 1, etc
Case 1
    MsgBox "Save", vbOKOnly + vbInformation, "You Chose..."
Case 2
    MsgBox "Print", vbOKOnly + vbInformation, "You Chose..."
Case 3
    MsgBox "Find...", vbOKOnly + vbInformation, "You Chose..."
Case 4
    MsgBox "Close", vbOKOnly + vbInformation, "You Chose..."
    'set variable to return a value
    Ret = MsgBox("Do you want to Close?", vbYesNo + vbInformation, "Quit?!?")
    'if returned value is vbyes
    If Ret = vbYes Then
    'unload form
        Unload Me
    'end if
    End If
'End of evaluation
End Select
End Sub

Private Sub mnuItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Condition
If mnuItem(Index).Top < mnuItem(Index).Tag + 1 Then
    'Set mnuItem top property +4
    mnuItem(Index).Top = mnuItem(Index).Top + 4
'end if condition
End If
End Sub

Private Sub mnuItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Sets label forecolor back to black
    mnuItem(Index).ForeColor = vbBlue
End Sub

Private Sub mnuItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Condition
If mnuItem(Index).Top < mnuItem(Index).Tag + 1 Then
    'Set mnuItem top property -4
    mnuItem(Index).Top = mnuItem(Index).Top - 4
'end if condition
End If
End Sub

Private Sub picMenu_Click()
'call another sub
    Call Form_Click
End Sub

Private Sub picMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Sets label forecolor back to black
    lblFile.ForeColor = vbBlack
'for loop
For A = 0 To mnuItem.Count - 1
'Sets label forecolor back to black
    mnuItem(A).ForeColor = vbBlack
'continues for loop
Next A
End Sub

Private Sub picMnuFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'for loop
For A = 0 To mnuItem.Count - 1
'Sets label forecolor back to black
    mnuItem(A).ForeColor = vbBlack
'continues for loop
Next A
End Sub

Private Sub tmrMenu_Timer()
For I = Opacity To 255 Step 30
If I > 254 - 30 Then
    'hide the picture box
    picMnuFile.Visible = False
    'Disable the timer
    tmrMenu.Enabled = False
End If
'Clear picture box
    picMnuFile.Cls
With BF
    .BlendOp = AC_SRC_OVER
    .BlendFlags = 0
    .SourceConstantAlpha = I
    .AlphaFormat = 0
End With
'api calls
    RtlMoveMemory lBF, BF, 4
    AlphaBlend picMnuFile.hdc, 0, 0, picMnuFile.ScaleWidth, picMnuFile.ScaleHeight, Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, lBF
'refresh picture box contents
    picMnuFile.Refresh
Next I
End Sub

Private Sub tmrMenuItem_Timer()
For I = 255 To Opacity Step -30
'disable timer
    tmrMenuItem.Enabled = False
'Clear picture box
    picMnuFile.Cls
With BF
    .BlendOp = AC_SRC_OVER
    .BlendFlags = 0
    .SourceConstantAlpha = I
    .AlphaFormat = 0
End With
    'api calls
    RtlMoveMemory lBF, BF, 4
    AlphaBlend picMnuFile.hdc, 0, 0, picMnuFile.ScaleWidth, picMnuFile.ScaleHeight, Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, lBF
    'Refresh Picture Box
    picMnuFile.Refresh
Next I
End Sub

Private Sub txtOpacity_Click()
'call sub
    Call Form_Click
End Sub

Private Sub txtOpacity_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'call sub
    Call Form_MouseMove(Button, Shift, X, Y)
End Sub

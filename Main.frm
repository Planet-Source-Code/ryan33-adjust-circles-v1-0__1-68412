VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Adjustable Circles"
   ClientHeight    =   2325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrMoveDown 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3480
      Top             =   1680
   End
   Begin VB.Timer tmrMoveUp 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3480
      Top             =   720
   End
   Begin VB.Timer tmrMoveRight 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3960
      Top             =   1200
   End
   Begin VB.Timer tmrMoveLeft 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3000
      Top             =   1200
   End
   Begin VB.Timer tmrDrawCircleSmaller 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3720
      Top             =   120
   End
   Begin VB.Timer tmrDrawCircleBigger 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3120
      Top             =   120
   End
   Begin VB.Label lblTitleHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "Adjust Circles Program (Click for Help)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblCircleSize 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Circle Size: ? (None)"
      ForeColor       =   &H00C0C000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblCurrentKey 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Key: ? (None)"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim CircleSize, DefaultSize, MaxSize, SizeChange, MoveSpeed, GoAmmount As Integer


Public Sub Form_Load()

    ' Set the defaults at startup (can play with the values below to get different results)
    CircleSize = 1000
    DefaultSize = 50
    SizeChange = 20
    MoveSpeed = 10
    GoAmmount = 40
    MaxSize = Screen.Height / 2 - 400 ' middle of the current screen
    
    'Me.Refresh
    'Me.Circle (Screen.Width / 2, Screen.Height / 2), CircleSize, Me.FillColor
    
    lblTitleHelp_Click
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    ' Make the circle bigger as user holds down the A key
    If KeyCode = vbKeyA Then tmrDrawCircleBigger.Enabled = True
    
    ' Make the circle smaller as user holds down the Z key
    If KeyCode = vbKeyZ Then tmrDrawCircleSmaller.Enabled = True
    
    ' Move Up
    If KeyCode = vbKeyUp Then tmrMoveUp.Enabled = True
    ' Move Left
    If KeyCode = vbKeyLeft Then tmrMoveLeft.Enabled = True
    ' Move Right
    If KeyCode = vbKeyRight Then tmrMoveRight.Enabled = True
    ' Move Down
    If KeyCode = vbKeyDown Then tmrMoveDown.Enabled = True
    
    ' Exit if user presses the ESC key
    If KeyCode = vbKeyEscape Then Unload Me
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    ' Disable keys when not pressed
    If KeyCode = vbKeyA Then tmrDrawCircleBigger.Enabled = False
    If KeyCode = vbKeyZ Then tmrDrawCircleSmaller.Enabled = False
    
    If KeyCode = vbKeyUp Then tmrMoveUp.Enabled = False
    If KeyCode = vbKeyLeft Then tmrMoveLeft.Enabled = False
    If KeyCode = vbKeyRight Then tmrMoveRight.Enabled = False
    If KeyCode = vbKeyDown Then tmrMoveDown.Enabled = False
    
    ' Show current key in use to the user
    lblCurrentKey.Caption = "Current Key: ? (None)"
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Creates a circle (also clears teh form at same time)
    If Button = vbRightButton Then Me.Circle (X, Y), CircleSize, Me.FillColor
    
    ' Just clears the form
    If Button = vbMiddleButton Then Me.Cls
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Drag the circle around with the left mouse button (draw)
    If Button = vbLeftButton Then Me.Circle (X, Y), CircleSize, Me.FillColor
    
    ' Just drag and dont draw
    If Button = vbRightButton Then
        Me.Refresh
        Me.Circle (X, Y), CircleSize, Me.FillColor
    End If
    
End Sub

Private Sub Form_Resize()
    ' Center the text
    'lblTitleHelp.Left = Me.Width \ 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Hide
    MsgBox "Thanks for checking out this small program I made. If you really like it then please vote for me and leave comments. If I get feedback from this (or any of my other programs i made) then I will make better improved versions of them. Thanks! :)" + vbNewLine + vbNewLine + "'AdjustCircles' Written by: Ryan27", vbInformation, "Credits"
End Sub

Private Sub lblTitleHelp_Click()
    MsgBox "Controls:" + vbNewLine + vbNewLine + _
    "Press A to Increase circle size and Z to decrease." + vbNewLine + _
    "Press UP,DOWN, LEFT, AND RIGHT arrows to move the circle around." + vbNewLine + _
    "Move the mouse around while holding the left mouse button to draw." + vbNewLine + _
    "Press and release (without moving the mouse) to make circles." + vbNewLine + _
    "Middle Mouse button to clear the form." + vbNewLine + _
    "Press ESC to exit when your done." + vbNewLine + vbNewLine + _
    "Enjoy! :)", vbInformation, "Controls"
End Sub

Public Sub tmrDrawCircleBigger_Timer()

    ' Print to user
    lblCurrentKey.Caption = "Current Key: A (Bigger)"
    lblCircleSize.Caption = "Current Circle Size: " & CircleSize
    
    ' How much to increase the size of the circle each time
    CircleSize = CircleSize + SizeChange
    
    Me.Refresh
    ' For Reference: Me.Circle (X, Y), Z, Color
    ' X = side to side, and Y = up/down, Z = in/out or size
    Me.Circle (Screen.Width / 2, Screen.Height / 2), CircleSize, Me.FillColor
    
    ' Make sure it dont get too big so you dont get a over flow error
    If CircleSize >= MaxSize Then
        CircleSize = MaxSize
        tmrDrawCircleBigger.Enabled = False
        Exit Sub
    End If
        
End Sub

Public Sub tmrDrawCircleSmaller_Timer()
    
    ' Print to user
    lblCurrentKey.Caption = "Current Key: Z (Smaller)"
    lblCircleSize.Caption = "Current Circle Size: " & CircleSize
    
    ' How much to increase the size of the circle each time
    CircleSize = CircleSize - SizeChange
    
    Me.Refresh
    ' For Reference: Me.Circle (X, Y), Z, Color
    ' X = side to side, and Y = up/down, Z = in/out or size
    Me.Circle (Screen.Width / 2, Screen.Height / 2), CircleSize, Me.FillColor
    
    ' Make sure it dont get too big so you dont get a over flow error
    If CircleSize <= DefaultSize Then
        CircleSize = DefaultSize
        tmrDrawCircleSmaller.Enabled = False
        Exit Sub
    End If
    
End Sub

Public Sub tmrMoveLeft_Timer()

    lblCurrentKey.Caption = "Current Key: Left Arrow"
    
    ' Move circle to the LEFT slowly like a video game character
    MoveSpeed = MoveSpeed - GoAmmount
    Me.Refresh
    ' For Reference: Me.Circle (X, Y), Z, Color
    ' X = side to side, and Y = up/down, Z = in/out or size
    Me.Circle (Screen.Width / 2 + MoveSpeed, Screen.Height / 2), CircleSize, Me.FillColor
    
End Sub

Public Sub tmrMoveRight_Timer()

    lblCurrentKey.Caption = "Current Key: Right Arrow"
    
    ' Move circle to the RIGHT slowly like a video game character
    MoveSpeed = MoveSpeed + GoAmmount
    Me.Refresh
    ' For Reference: Me.Circle (X, Y), Z, Color
    ' X = side to side, and Y = up/down, Z = in/out or size
    Me.Circle (Screen.Width / 2 + MoveSpeed, Screen.Height / 2), CircleSize, Me.FillColor
    
End Sub

Public Sub tmrMoveUp_Timer()

    lblCurrentKey.Caption = "Current Key: Up Arrow"
    
    ' Move circle to the UP slowly like a video game character
    MoveSpeed = MoveSpeed - GoAmmount
    Me.Refresh
    ' For Reference: Me.Circle (X, Y), Z, Color
    ' X = side to side, and Y = up/down, Z = in/out or size
    Me.Circle (Screen.Width / 2, Screen.Height / 2 + MoveSpeed), CircleSize, Me.FillColor
        
End Sub

Private Sub tmrMoveDown_Timer()
    
    lblCurrentKey.Caption = "Current Key: Down Arrow"
    
     ' Move circle to the UP slowly like a video game character
    MoveSpeed = MoveSpeed + GoAmmount
    Me.Refresh
    ' For Reference: Me.Circle (X, Y), Z, Color
    ' X = side to side, and Y = up/down, Z = in/out or size
    Me.Circle (Screen.Width / 2, Screen.Height / 2 + MoveSpeed), CircleSize, Me.FillColor
    
End Sub

Private Sub Form_Click()
    'Unload Me
End Sub


VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   ScaleHeight     =   576
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   826
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   1620
      Picture         =   "Snake Picture.frx":0000
      ScaleHeight     =   1455
      ScaleWidth      =   5370
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   5370
   End
   Begin VB.Timer Timer1 
      Left            =   8700
      Top             =   5460
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Begin"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   2115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hdcDest As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal xSource As Long, ByVal ySource As Long, ByVal RasterOp As Long) As Long

Private Type Sections
    PosX As Variant
    PosY As Variant
    MoveX As Variant
    MoveY As Variant
    DirX As Variant
    DirY As Variant
    Step As Variant
    MaxY As Variant
    MaxX As Variant
    OX As Variant
    OY As Variant
    YActive As Boolean
    XActive As Boolean
    YStop As Boolean
    XStop As Boolean
End Type

Private PStop As Boolean
Private Sects() As Sections
Private LastYDir As Single
Private LastXDir As Single
Private AllowNext As Boolean
Private LColour As Long
Private NoMoreHits As Boolean

Private Sub Command1_Click()
If PStop = True And AllowNext = True And NoMoreHits = False Then
    PStop = False
    AllowNext = False
    Command1.Caption = "Stop This Sharade"
    StartSharade
Else
    PStop = True
    Command1.Caption = "Stopping, Please Wait"
    NoMoreHits = True
End If

End Sub

Private Sub StartSharade()


Dim SPosX As Single, SPosY As Single
Dim LP As Single

SPosX = Picture1.Top
SPosY = Picture1.Left
Dim MM As Variant

If (Screen.Width / 15) = 800 Then MM = 4
If (Screen.Width / 15) = 1024 Then MM = 5
If (Screen.Width / 15) = 1280 Then MM = 6

If MM = "" Then MM = 5

For LP = 0 To UBound(Sects)
    Sects(LP).PosX = SPosX
    Sects(LP).PosY = SPosY + (UBound(Sects) - LP)
    Sects(LP).DirX = 1
    Sects(LP).DirY = 1
    Sects(LP).Step = 0.05
    Sects(LP).MaxY = MM
    Sects(LP).MaxX = MM
    Sects(LP).YActive = False
    Sects(LP).XActive = False
    Sects(LP).MoveX = 0
    Sects(LP).MoveY = 0
    
    Next LP
    
    Sects(0).YActive = True
    Sects(0).XActive = True
    Timer1.Interval = 1
    
    
End Sub

Private Sub Command2_Click()
PStop = True
End Sub

Private Sub Form_Activate()
    ReDim Sects(Picture1.Height)
    PStop = True
    AllowNext = True
End Sub

Private Sub Timer1_Timer()
Dim LP As Single
Dim Over As Single

Form1.Cls
For LP = UBound(Sects) To 1 Step -1
    If Sects(LP - 1).PosY > Sects(LP).PosY + 4 And Sects(LP).YStop = False Then
        Sects(LP).YActive = True
    End If
    
    If Sects(LP - 1).PosX > Sects(LP).PosX + 4 And Sects(LP).XStop = False Then
        Sects(LP).XActive = True
    End If
Next LP

For LP = 0 To UBound(Sects)
    If Sects(LP).YActive = True Then
        
        
        
        If Sects(LP).DirY = 1 Then
            Sects(LP).MoveY = Sects(LP).MoveY + Sects(LP).Step
        ElseIf Sects(LP).DirY = 2 Then
            Sects(LP).MoveY = Sects(LP).MoveY - Sects(LP).Step
        End If
        
        
        
            If PStop = True Then
                If Sects(LastYDir).MoveY < 0.5 And Sects(LastYDir).MoveY > -0.5 And Sects(LastYDir).DirY = 1 Then
                    Sects(LastYDir).YStop = True
                    LastYDir = LastYDir + 1
                    If LastYDir = UBound(Sects) + 1 Then LastYDir = UBound(Sects)
                End If
            End If
            
            
            If Sects(LP).YStop = False Then
            If Sects(LP).MoveY > Sects(LP).MaxY Then
                Sects(LP).MoveY = Sects(LP).MaxY: Sects(LP).DirY = 2
            End If
            
            If Sects(LP).MoveY < Sects(LP).MaxY * -1 Then
                Sects(LP).DirY = 1
            End If
            
            
        
        
        Sects(LP).PosY = Sects(LP).PosY + Sects(LP).MoveY
        End If
        If Sects(LP).PosY > Int(Screen.Height / 15) Then
            Over = Sects(LP).PosY - Int(Screen.Height / 15)
            Sects(LP).PosY = Over
        End If
        
        If Sects(LP).PosY < 0 Then
            Over = Sects(LP).PosY
            Sects(LP).PosY = (Screen.Height / 15) - (Over * -1)
        End If
        
    End If
    
    
    
    
    If Sects(LP).XActive = True Then
        
        If Sects(LP).DirX = 1 Then
            Sects(LP).MoveX = Sects(LP).MoveX + Sects(LP).Step
        ElseIf Sects(LP).DirX = 2 Then
            Sects(LP).MoveX = Sects(LP).MoveX - Sects(LP).Step
        End If
        
        
            If PStop = True Then
                If Sects(LastXDir).MoveX < 0.5 And Sects(LastXDir).MoveX > -0.5 And Sects(LastXDir).DirX = 1 Then
                    Sects(LastXDir).XStop = True
                    LastXDir = LastXDir + 1
                    If LastXDir = UBound(Sects) + 1 Then LastXDir = UBound(Sects)
                End If
            End If
        
        
        If Sects(LP).XStop = False Then
        If Sects(LP).MoveX > Sects(LP).MaxX Then Sects(LP).MoveX = Sects(LP).MaxX: Sects(LP).DirX = 2
        If Sects(LP).MoveX < Sects(LP).MaxX * -1 Then
            Sects(LP).DirX = 1
            
        End If
        
        Sects(LP).PosX = Sects(LP).PosX + Sects(LP).MoveX
        End If
        
        If Sects(LP).PosX > Int(Screen.Width / 15) Then
            Over = Sects(LP).PosX - Int(Screen.Width / 15)
            Sects(LP).PosX = Over
        End If
        
        If Sects(LP).PosX < 0 Then
            Over = Sects(LP).PosX
            Sects(LP).PosX = (Screen.Width / 15) - (Over * -1)
        End If
        
    End If
    
    
    If Sects(UBound(Sects)).XStop = True And Sects(UBound(Sects)).YStop = True Then
        Timer1.Interval = 0
        MsgBox "Tadaaaaaaa."
        End
    End If
        
        BitBlt Form1.hDC, CLng(Sects(LP).PosX), CLng(Sects(LP).PosY), Picture1.Width, 1, Picture1.hDC, 1, Picture1.Height - LP, vbSrcCopy
    
Next LP
DoEvents

End Sub



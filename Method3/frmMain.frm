VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RTB Coloring - Method 3"
   ClientHeight    =   8055
   ClientLeft      =   2745
   ClientTop       =   1620
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   10455
   Begin RichTextLib.RichTextBox RTB 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   14208
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bDirty As Boolean

Private Sub Form_Load()

    ' setup and load!!
    InitKeyWords
    LoadFile RTB, App.Path & "\text.txt"
        
End Sub

Private Sub RTB_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lCursor As Long
Dim lSelectLen As Long
Dim lStart As Long
Dim lFinish As Long
Dim sText As String

    ' ------------------------------
    ' here's the on the fly coloring
    ' ------------------------------
    
    ' check for Ctrl+C
    If KeyCode = vbKeyC And Shift = 2 Then Exit Sub
    
    ' check for text being pasted into the box
    If KeyCode = vbKeyV And Shift = 2 Then
        
        Screen.MousePointer = vbHourglass
        DoClipBoardPaste RTB
        KeyCode = 0
        Screen.MousePointer = vbNormal
        Exit Sub
        
    End If
    
    ' if the cursor is moving to a different
    ' line then process the orginal line
    If KeyCode = 13 Or _
         KeyCode = vbKeyUp Or _
            KeyCode = vbKeyDown Then
    
        ' only color this line if it's been changed
        If bDirty Or KeyCode = 13 Then
                        
            ' lock the window to cancel out flickering
            LockWindowUpdate RTB.hWnd
            
            ' store the current cursor pos
            ' and current selection if there is any
            lCursor = RTB.SelStart
            lSelectLen = RTB.SelLength
            
            ' get the line start and end
            If lCursor <> 0 Then
                lStart = InStrRev(RTB.Text, vbCrLf, RTB.SelStart) + 2
                If lStart = 2 Then lStart = 1
            Else
                lStart = 1
            End If
            lFinish = InStr(lCursor + 1, RTB.Text, vbCrLf)
            If lFinish = 0 Then lFinish = Len(RTB.Text)
            
            ' do the coloring
            basColor.sText = RTB.Text
            DoColor RTB, lStart, lFinish
            
            ' if ENTER was pressed, we should color the next line
            ' as well, so that if a line is broken by the ENTER
            ' the new line and the old line are colored properly
            If KeyCode = 13 Then
                lStart = lCursor + 1
                lFinish = InStr(lStart, RTB.Text, vbCrLf)
                If lFinish = 0 Then lFinish = Len(RTB.Text)
                ' only color if another line exists
                If lStart - 1 <> lFinish Then
                  RTB.SelStart = lStart - 1
                  RTB.SelLength = lFinish - lStart
                  RTB.SelColor = vbBlack
                  DoColor RTB, lStart, lFinish
               End If
            End If
            
            ' reset the properties
            RTB.SelStart = lCursor
            RTB.SelLength = lSelectLen
            RTB.SelColor = vbBlack
            
            ' reset the flag and release the window
            bDirty = False
            LockWindowUpdate 0&
            
        End If
        
    ElseIf Not IsControlKey(KeyCode) Then
                
        ' a different key was pressed - and
        ' this will alter the line so it
        ' needs recoloring when we move off it
        If Not bDirty Then
            
            LockWindowUpdate RTB.hWnd
            
            ' get the line start and end
            lStart = InStrRev(RTB.Text, vbCrLf, RTB.SelStart + 1) + 1
            lFinish = InStr(RTB.SelStart + 1, RTB.Text, vbCrLf)
            If lFinish = 0 Then lFinish = Len(RTB.Text)
            
            ' color the line (remembering the cursor position)
            lCursor = RTB.SelStart
            lSelectLen = RTB.SelLength
            RTB.SelStart = lStart
            RTB.SelLength = lFinish - lStart
            RTB.SelColor = vbBlack
            RTB.SelStart = lCursor
            RTB.SelLength = lSelectLen
            bDirty = True
            
            LockWindowUpdate 0&
            
        End If
        
    End If
End Sub

Private Function IsControlKey(ByVal KeyCode As Long) As Boolean
    ' check if the key is a control key
    Select Case KeyCode
        Case vbKeyLeft, vbKeyRight, vbKeyHome, _
             vbKeyEnd, vbKeyPageUp, vbKeyPageDown, _
             vbKeyShift, vbKeyControl
            IsControlKey = True
        Case Else
            IsControlKey = False
    End Select
End Function


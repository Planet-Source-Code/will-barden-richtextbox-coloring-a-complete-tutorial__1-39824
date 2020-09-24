VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RTB Coloring - Method 2"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RTB 
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
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

Private Sub Form_Load()
Dim lStart As Long
Dim lFinish As Long
Dim Text As String

    LoadFile
    InitWords
    
    ' split the text into lines and color them one by one
    Text = RTB.Text
    lStart = 1
    Do While lStart <> 2 And lStart < Len(Text)
        ' find the end of this line
        lFinish = InStr(lStart + 1, Text, vbCrLf)
        If lFinish = 0 Then lFinish = Len(Text)
            
        ' color it
        DoColor RTB, lStart, lFinish
        
        ' move up to get the next line
        lStart = lFinish + 2
    Loop
    
    ' reset the cursor
    RTB.SelStart = 0
End Sub

Private Sub LoadFile()
Dim FF As Long
    FF = FreeFile
    Open App.Path & "\text.txt" For Input As FF
        RTB.Text = Input(LOF(FF), FF)
    Close FF
End Sub

Private Sub RTB_KeyDown(KeyCode As Integer, Shift As Integer)
Static bDirty As Boolean
Dim lCursor As Long
Dim lSelectLen As Long
Dim lStart As Long
Dim lFinish As Long

    '
    ' here's the on the fly coloring
    '
    
    ' if the cursor is moving to a different line
    ' then process the orginal line
    If KeyCode = 13 Or _
        KeyCode = vbKeyUp Or _
            KeyCode = vbKeyDown Then
    
        ' only color this line if it's been changed
        If bDirty Then
            
            ' store the current cursor pos
            ' and current selection if there is any
            lCursor = RTB.SelStart
            lSelectLen = RTB.SelLength
            
            ' get the line start and end
            lStart = InStrRev(RTB.Text, vbCrLf, RTB.SelStart + 1) + 2
            lFinish = InStr(RTB.SelStart + 1, RTB.Text, vbCrLf)
            If lFinish = 0 Then lFinish = Len(RTB.Text)
            
            ' do the coloring
            DoColor RTB, lStart, lFinish
            
            ' reset the cursor
            RTB.SelStart = lCursor
            RTB.SelLength = lSelectLen
            
            ' restore the color
            RTB.SelColor = vbBlack
            
            ' reset the flag
            bDirty = False
            
        End If
        
    ElseIf KeyCode <> vbKeyLeft And _
        KeyCode <> vbKeyRight And _
            KeyCode <> vbKeyHome And _
                KeyCode <> vbKeyEnd And _
                    KeyCode <> vbKeyPageUp And _
                        KeyCode <> vbKeyPageDown Then
                
        ' a different key was pressed - and
        ' this will alter the line so it
        ' needs recoloring when we move off it
        If Not bDirty Then
                
            ' remove all coloring from this line
            ' as a visual reminder to the user
            ' that it has been changed
            
            ' get the line start and end
            lStart = InStrRev(RTB.Text, vbCrLf, RTB.SelStart + 1) + 2
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
            
        End If
        
    End If
End Sub

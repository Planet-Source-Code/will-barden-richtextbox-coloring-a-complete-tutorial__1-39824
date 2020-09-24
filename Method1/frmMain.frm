VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RTB Coloring - Method 1"
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
      Enabled         =   -1  'True
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
    LoadFile
    InitWords
    DoColor RTB
    RTB.SelStart = 0
End Sub

Private Sub LoadFile()
Dim FF As Long
    FF = FreeFile
    Open App.Path & "\text.txt" For Input As FF
        RTB.Text = Input(LOF(FF), FF)
    Close FF
End Sub

Private Sub RTB_KeyUp(KeyCode As Integer, Shift As Integer)
Dim lCursor As Long
    '
    ' here's the on the fly coloring
    '
    
    ' store the current cursor pos
    lCursor = RTB.SelStart
    
    ' do the coloring
    DoColor RTB
    
    ' reset the cursor
    RTB.SelStart = lCursor
    
    ' restore the color
    RTB.SelColor = vbBlack
End Sub

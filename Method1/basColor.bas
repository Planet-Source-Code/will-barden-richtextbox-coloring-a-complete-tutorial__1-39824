Attribute VB_Name = "basColor"
Option Explicit

Type WORD_TYPE
    Text As String  ' word to be colored
    Color As Long   ' color to color the word :)
End Type

Dim Words() As WORD_TYPE

Public Sub InitWords()
    ' initialize the array of words
    ReDim Words(0 To 10)
    Words(0).Text = "Option"
    Words(0).Color = vbBlue
    Words(1).Text = "Explicit"
    Words(1).Color = vbBlue
    Words(2).Text = "Type"
    Words(2).Color = vbBlue
    Words(3).Text = "As"
    Words(3).Color = vbBlue
    Words(4).Text = "String"
    Words(4).Color = vbBlue
    Words(5).Text = "End"
    Words(5).Color = vbBlue
    Words(6).Text = "Dim"
    Words(6).Color = vbBlue
    Words(7).Text = "ReDim"
    Words(7).Color = vbBlue
    Words(8).Text = "Public"
    Words(8).Color = vbBlue
    Words(9).Text = "Sub"
    Words(9).Color = vbBlue
    Words(10).Text = "ByVal"
    Words(10).Color = vbBlue
End Sub

Public Sub DoColor(RTB As RichTextBox)
Dim i As Long
Dim p1 As Long, p2 As Long
Dim Text As String
Dim sTmp As String
    
    ' cache the text - speeds things up a bit
    Text = RTB.Text
    
    ' go through each item in the Words array
    For i = LBound(Words) To UBound(Words)
    
        ' find each instance of the word in the rtb
        p1 = InStr(1, Text, Words(i).Text)
        Do While p1 > 0
        
            ' color it to the appropriate color
            RTB.SelStart = p1 - 1
            RTB.SelLength = Len(Words(i).Text)
            RTB.SelColor = Words(i).Color
            
            ' go on to the next word
            p1 = InStr(p1 + 1, Text, Words(i).Text)

        Loop

    Next i
    
    ' go through and color all the comment lines
    p1 = 1
    Do While p1 <> 2 And p1 < Len(Text)
        
        ' find the next eol character
        p2 = InStr(p1 + 1, Text, vbCrLf)
        If p2 = 0 Then p2 = Len(Text)
        
        ' grab this line into a temp variable
        sTmp = Mid$(Text, p1, p2 - p1)
        
        ' if it's a comment line - color it
        If Left(Trim$(sTmp), 1) = "'" Then
            RTB.SelStart = p1
            RTB.SelLength = p2 - p1
            RTB.SelColor = vbRed
        End If
        
        ' move onto the next line
        p1 = p2 + 2
        
    Loop
    
End Sub

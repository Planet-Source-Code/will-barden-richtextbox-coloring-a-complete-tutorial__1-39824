Attribute VB_Name = "basColor"
Option Explicit

Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Const COL_KEYWORD = &HC00000    ' dark blue
Const COL_COMMENT = &H8000&     ' middle green

Type WORD_TYPE
    Text As String  ' word to be colored
    Color As Long   ' color to color the word :)
End Type

Dim Words() As WORD_TYPE

Public Sub InitWords()
    ' initialize the array of words
    ReDim Words(0 To 14)
    Words(0).Text = "Option"
    Words(0).Color = COL_KEYWORD
    Words(1).Text = "Explicit"
    Words(1).Color = COL_KEYWORD
    Words(2).Text = "Type"
    Words(2).Color = COL_KEYWORD
    Words(3).Text = "As"
    Words(3).Color = COL_KEYWORD
    Words(4).Text = "String"
    Words(4).Color = COL_KEYWORD
    Words(5).Text = "End"
    Words(5).Color = COL_KEYWORD
    Words(6).Text = "Dim"
    Words(6).Color = COL_KEYWORD
    Words(7).Text = "ReDim"
    Words(7).Color = COL_KEYWORD
    Words(8).Text = "Public"
    Words(8).Color = COL_KEYWORD
    Words(9).Text = "Sub"
    Words(9).Color = COL_KEYWORD
    Words(10).Text = "ByVal"
    Words(10).Color = COL_KEYWORD
    Words(11).Text = "If"
    Words(11).Color = COL_KEYWORD
    Words(12).Text = "Then"
    Words(12).Color = COL_KEYWORD
    Words(13).Text = "Else"
    Words(13).Color = COL_KEYWORD
    Words(14).Text = "Private"
    Words(14).Color = COL_KEYWORD
    
End Sub

Public Sub DoColor(RTB As RichTextBox, ByVal Start As Long, ByVal Finish As Long)
Dim i As Long
Dim p1 As Long, p2 As Long
Dim Text As String
Dim sLine As String
    
    ' cache the text - speeds things up a bit
    Text = RTB.Text
    
    ' clear the coloring
    RTB.SelStart = Start
    RTB.SelLength = Finish - Start
    RTB.SelColor = vbBlack
    
    ' check if this line is a comment line
    sLine = StripCRLF$(Trim$(Mid$(Text, Start, Finish - Start)))
    If Left$(sLine, 1) = "'" Then
        
        ' color the whole line green
        RTB.SelStart = Start - 1
        RTB.SelLength = Finish - Start
        RTB.SelColor = COL_COMMENT

    Else
    
        ' go through each item in the Words array
        For i = LBound(Words) To UBound(Words)
        
            ' find each instance of the word
            ' within the specified range
            p1 = InStr(Start, Text, Words(i).Text)
            Do While p1 > 0 And p1 < Finish
            
                ' color it to the appropriate color
                RTB.SelStart = p1 - 1
                RTB.SelLength = Len(Words(i).Text)
                RTB.SelColor = Words(i).Color
                
                ' go on to the next word
                p1 = InStr(p1 + 1, Text, Words(i).Text)
            Loop
            
        Next i
    
    End If
    
End Sub

Private Function StripCRLF$(ByVal sText As String)
    If Len(sText) > 1 Then
        If Mid$(sText, 1, 2) = vbCrLf Then
            sText = Mid$(sText, 3, Len(sText))
        End If
        If Len(sText) > 1 Then
            If Mid$(sText, Len(sText) - 2, 2) = vbCrLf Then
                sText = Mid$(sText, 1, Len(sText) - 2)
            End If
        End If
    End If
    StripCRLF = sText
End Function

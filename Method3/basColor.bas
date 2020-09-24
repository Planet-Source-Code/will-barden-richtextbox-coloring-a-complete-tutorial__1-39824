Attribute VB_Name = "basColor"
Option Explicit

'#--------------------------------------------------------------------------#
'#                                                                          #
'#    File..........: basColor [Method 3]                                   #
'#    Author........: Will Barden                                           #
'#    Last Modified.: 10/9/02                                               #
'#    Dependancies..: None                                                  #
'#                                                                          #
'#    This will do it all.. store the keywords, store the index             #
'#    sort and search them, then use them to color the contents             #
'#    of a specified RTB. Cool huh?                                         #
'#                                                                          #
'#    In order to use the module, first call InitKeyWords, then             #
'#    call LoadFile to fill and color the RTB. A little handling            #
'#    in the RTB_KeyDown event and you can call DoClipBoardPaste            #
'#    as well.                                                              #
'#                                                                          #
'#    At present this module colors at a rate of 90KB/sec on my             #
'#    700mHz/128MB RAM machine. Not bad I reckon.. ;) The keyword           #
'#    List is pitifully small, but I hope that enlarging it won't           #
'#    degrade performance too much.                                         #
'#                                                                          #
'#    Known bugs: If a line contains more than one of the same key          #
'#    word, often only the first few are colored.                           #
'#    Text that is replaced via Undo (Ctrl+Z) is not properly colored.      #
'#                                                                          #
'#--------------------------------------------------------------------------#

'#--------------------------------------------------------------------------#
'#  apis, enums, consts, declares
'#--------------------------------------------------------------------------#
' api to stop the window refreshing
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Const COL_KEYWORD = &H800000    ' dark blue
Const COL_COMMENT = &H8000&     ' middle green
Const CHAR_COMMENT = "'"        ' comment line char

Type WORD_TYPE
    Text As String  ' word to be colored
    Color As Long   ' color to color the word :)
End Type

Type LETTER_TYPE
    Start As Long   ' first time the letter appears in the list
    Finish As Long  ' last time the letter appears in the list
End Type

'#--------------------------------------------------------------------------#
'#  variables
'#--------------------------------------------------------------------------#
Dim Words() As WORD_TYPE
Dim Letters() As LETTER_TYPE
Dim Strings() As String
Public sText As String

'#--------------------------------------------------------------------------#
'#  methods
'#--------------------------------------------------------------------------#

'//--[InitKeyWords]-----------------------------------------------------------//
'
'  Builds the arrays of keywords, then builds
'  an alphabetical index of the array to aid
'  searching later on.
'
Public Sub InitKeyWords()
    ' initialize the array of words
    ReDim Words(0 To 36)
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
    Words(15).Text = "For"
    Words(15).Color = COL_KEYWORD
    Words(16).Text = "Next"
    Words(16).Color = COL_KEYWORD
    Words(17).Text = "To"
    Words(17).Color = COL_KEYWORD
    Words(18).Text = "Exit"
    Words(18).Color = COL_KEYWORD
    Words(19).Text = "Do"
    Words(19).Color = COL_KEYWORD
    Words(20).Text = "Loop"
    Words(20).Color = COL_KEYWORD
    Words(21).Text = "While"
    Words(21).Color = COL_KEYWORD
    Words(22).Text = "Until"
    Words(22).Color = COL_KEYWORD
    Words(23).Text = "DoEvents"
    Words(23).Color = COL_KEYWORD
    Words(24).Text = "Long"
    Words(24).Color = COL_KEYWORD
    Words(25).Text = "Byte"
    Words(25).Color = COL_KEYWORD
    Words(26).Text = "Single"
    Words(26).Color = COL_KEYWORD
    Words(27).Text = "Double"
    Words(27).Color = COL_KEYWORD
    Words(28).Text = "Integer"
    Words(28).Color = COL_KEYWORD
    Words(29).Text = "Function"
    Words(29).Color = COL_KEYWORD
    Words(30).Text = "And"
    Words(30).Color = COL_KEYWORD
    Words(31).Text = "Event"
    Words(31).Color = COL_KEYWORD
    Words(32).Text = "LBound"
    Words(32).Color = COL_KEYWORD
    Words(33).Text = "UBound"
    Words(33).Color = COL_KEYWORD
    Words(34).Text = "Xor"
    Words(34).Color = COL_KEYWORD
    Words(35).Text = "Const"
    Words(35).Color = COL_KEYWORD
    Words(36).Text = "Boolean"
    Words(36).Color = COL_KEYWORD
    
    ' sort the array
    CombSort Words
    
    ' build the index of letter positions
    BuildIndex
End Sub

'//--[LoadFile]-------------------------------------------------------------//
'
'  Loads and colors a file in the RTB
'
Public Sub LoadFile(RTB As RichTextBox, ByVal sFilePath As String)
Dim FF As Long
Dim lStart As Long
Dim lFinish As Long
Dim Text As String

    ' load the file
    FF = FreeFile
    Open sFilePath For Input As FF
        RTB.Text = Input(LOF(FF), FF)
    Close FF

    ' split the text into lines and color them one by one
    LockWindowUpdate RTB.hwnd
    RTB.Visible = False
    Text = RTB.Text
    basColor.sText = RTB.Text
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
    RTB.Visible = True
    LockWindowUpdate 0&

End Sub

'//--[DoColor]--------------------------------------------------------------//
'
'  Here it is - the beast itself. This routine colors
'  a single line of text within the RTB. It will
'  split each line up into words using the custom
'  split function (SplitWords), then match each word
'  against the list of keywords.
'
Public Sub DoColor(RTB As RichTextBox, ByVal lStart As Long, ByVal lFinish As Long)
Dim sWords()    As String
Dim sLine       As String
Dim sChar       As String
Dim lCurPos     As Long
Dim lIndex      As Long
Dim lColor      As Long
Dim lPos        As Long
Dim lPos2       As Long
Dim lCom        As Long
Dim i           As Long

    ' grab the line
    sLine = Trim$(Mid$(sText, lStart, lFinish - lStart))
    ' remove the EOL
    sLine = RemoveEOL(sLine)
    ' remove the quotes so they're not colored
    sLine = RemoveStrings(sLine)
    
    ' split the line into words using our custom function
    sWords = SplitWords(sLine)
    
    ' check each word against the list
    lCurPos = 1
    ' search for each word in the array
    For i = LBound(sWords) To UBound(sWords)
    
        If Trim$(sWords(i)) <> "" Then
    
            ' check for comment in the middle of a line
            If Left$(sWords(i), 1) = CHAR_COMMENT Then
            
                ' color the rest of the line
                RTB.SelStart = InStr(lStart, sText, sWords(i)) - 1
                RTB.SelLength = Len(sWords(i))
                RTB.SelColor = COL_COMMENT
            
            Else
        
                ' its a normal keyword - so color it
                ' first get the array positions from
                ' the index
                If sWords(i) = "ByVal" Or sWords(i) = "Byte" Then
                    DoEvents: End If
                sChar = Left$(LCase$(sWords(i)), 1)
                ' if we've got a valid alphabetic char
                If sChar <> "" Then
                    ' convert this char to an index in the letters array
                    lIndex = Asc(sChar) - 97
                    ' if the index is a valid one - this
                    ' means that the text is a word, so
                    ' we should try to color it
                    If lIndex >= 0 And lIndex < UBound(Letters) Then
                        ' color the word, passing the index parameters
                        lColor = GetColor(sWords(i), _
                                    Letters(lIndex).Start, _
                                    Letters(lIndex).Finish)
                        ' if a color was returned - color the word
                        If lColor Then
                            ' locate the word in the line
                            RTB.SelStart = InStr(lStart + lCurPos - 1, sText, sWords(i)) - 1
                            RTB.SelLength = Len(sWords(i))
                            RTB.SelColor = lColor
                        End If
                    End If
                End If ' sChar <> ""
            End If ' CHAR_COMMENT
        End If ' sWords(i) <> ""
        
        ' move the current position within the line on
        lCurPos = lCurPos + Len(sWords(i))
        
    Next i
        
End Sub

'//--[DoClipBoardPaste]-----------------------------------------------------//
'
'  Call this when text has been pasted into the
'  RTB. It will grab the text, split it into lines
'  and color it.
'
Public Sub DoClipBoardPaste(RTB As RichTextBox)
Dim lCursor As Long
Dim lStart As Long
Dim lFinish As Long
Dim sText As String
Dim p1 As Long, p2 As Long

    ' store the cursor position
    lCursor = RTB.SelStart
    
    ' add the text and color it
    LockWindowUpdate RTB.hwnd
    
    ' get the text to be pasted from the clipboard
    sText = Clipboard.GetText
    
    ' get the start point - this should be the previous
    ' vbCrLf to where the text was inserted, to make
    ' sure that if it's inserted mid-line, the whole
    ' line is colored
    lStart = InStrRev(RTB.Text, vbCrLf, RTB.SelStart) + 2
    If lStart = 0 Then lStart = RTB.SelStart
    ' also store the finish point
    lFinish = RTB.SelStart + Len(sText)
    
    ' now add the text to the box
    RTB.SelText = sText
    basColor.sText = RTB.Text
    
    ' now color each line individually starting
    ' from lStart since this is the position of
    ' the first changed line
    p1 = lStart
    Do
        ' find the next EOL character, this combined
        ' with lStart gives us the line to color
        p2 = InStr(p1, RTB.Text, vbCrLf)
        If p2 = 0 Then p2 = lFinish
                    
        ' now strip out this line and color it
        ' color it black first to remove any
        ' previous coloring..
        RTB.SelStart = p1 - 1
        RTB.SelLength = p2 - p1
        RTB.SelColor = vbBlack
        DoColor RTB, p1, p2
        
        ' move the start pointer on to just after
        ' the last EOL character - essentially onto
        ' the next actual line of text
        p1 = p2 + 2
              
        ' exit condition - keep going until we can't
        ' find any more vbCrLf (<>2) and while
        ' p1 (the start of line pointer) is lower
        ' that lFinish (the end of the text we're
        ' coloring)... easy enough
        If p1 = 2 Or p1 >= lFinish + 2 Then Exit Do
        DoEvents
    Loop
    
    ' restore the original values
    RTB.SelStart = lCursor + Len(sText)
    RTB.SelColor = vbBlack
    
    ' null the keypress (to avoid the text pasting twice)
    LockWindowUpdate 0&

End Sub

'#--------------------------------------------------------------------------#
'#  private internals
'#--------------------------------------------------------------------------#

'//--[BuildIndex]----------------------------------------------------------//
'
'  Takes the Words array and constructs an alphabetical
'  index which it puts into the Letters array.
'  Each item in the letters array accounts for a letter
'  in the alphabet - Letters(0) = "a".
'  The .Start property is the Index in the Words array
'  at which that letter starts, and the finish is the
'  same. The purpose of this is to get Hi and Lo params
'  for the GetColor (a standard binary search algorithm).
'  This saves several loops round the algorithm.
'
Private Sub BuildIndex()
Dim i As Long, j As Long
Dim sChar As String
Dim bStart As Boolean

    ' go through each letter in the alphabet
    ReDim Letters(25)
    For i = 0 To 25
        ' get the current char
        sChar = Chr$(i + 97)
        ' find the first and last instances of the letter
        For j = LBound(Words) To UBound(Words)
            If Left$(LCase$(Words(j).Text), 1) = sChar Then
                If Not bStart Then
                    ' found the start
                    bStart = True
                    Letters(i).Start = j
                End If
                ' if we've hit the end of the list
                If j = UBound(Words) Then
                    Letters(i).Finish = j
                    Exit Sub
                End If
            Else
                ' its a different char
                If bStart Then
                    ' we've found the end
                    Letters(i).Finish = j - 1
                    bStart = False
                    Exit For
                End If
                ' see if we've gone too far -
                ' there are no words beginning with
                ' this letter in the list
                If Left$(LCase$(Words(j).Text), 1) > sChar Then
                    Exit For
                End If
            End If
        Next j
    Next i

End Sub

'//--[GetColor]--------------------------------------------------------------//
'
'  Searches the Words array for a match using a standard
'  binary search algorithm, using the Lo and Hi params
'  as starting points.
'
Private Function GetColor(ByVal sWord As String, _
                          ByVal Lo As Long, _
                          ByVal Hi As Long) As Long
Dim lHi As Long
Dim lLo As Long
Dim lMid As Long
    
    ' standard binary search the words array
    ' return the color if a match is found
    lLo = Lo
    lHi = Hi
    Do While lHi >= lLo
        lMid = (lLo + lHi) \ 2
        If LCase$(Words(lMid).Text) = LCase$(sWord) Then
            GetColor = Words(lMid).Color
            Exit Do
        End If
        If LCase$(Words(lMid).Text) > LCase$(sWord) Then
            lHi = lMid - 1
        Else
            lLo = lMid + 1
        End If
    Loop
    
End Function

'//--[SplitWords]---------------------------------------------------------//
'
'  Since splitting a line into words by a single
'  character is not acceptable because we have to
'  take several end of word characters into account,
'  this routine was written.
'  It searches through the string from left to right
'  and locates the nearest word break char from a list
'  then splits at that word.
'
Private Function SplitWords(ByVal sText As String) As String()
Dim i As Long, lPos As Long
Dim sWords() As String
Dim sWordBreaks(0 To 8) As String
Dim lBreakPoints() As Long
Dim lBreak As Long
    
    ' list of word break characters
    sWordBreaks(0) = " "
    sWordBreaks(1) = "("
    sWordBreaks(2) = ")"
    sWordBreaks(3) = "<"
    sWordBreaks(4) = ">"
    sWordBreaks(5) = "."
    sWordBreaks(6) = ","
    sWordBreaks(7) = "="
    sWordBreaks(8) = CHAR_COMMENT ' comments
    ReDim lBreakPoints(UBound(sWordBreaks))

    ' get them words!
    ReDim sWords(0)
    lPos = 1
    Do
    
        ' locate the word break points
        For i = 0 To UBound(sWordBreaks)
            lBreakPoints(i) = InStr(lPos, sText, sWordBreaks(i))
        Next i
        
        ' now work out which is closest
        lBreak = Len(sText) + 1
        For i = 0 To UBound(lBreakPoints)
            If lBreakPoints(i) <> 0 Then
                If lBreakPoints(i) < lBreak Then lBreak = lBreakPoints(i)
            End If
        Next i
    
        ' now split out the word
        ' if no break point was found, then we've
        ' hit the end of the line, so add all the rest
        If lBreak = Len(sText) + 1 Then
            sWords(UBound(sWords)) = Mid$(sText, lPos)
        Else
            ' add this word - first check for a comment
            If Mid$(sText, lBreak, 1) = CHAR_COMMENT Then
                ' first add the word
                sWords(UBound(sWords)) = Mid$(sText, lPos, lBreak - lPos)
                ' then add the rest as a comment
                ReDim Preserve sWords(UBound(sWords) + 1)
                sWords(UBound(sWords)) = Mid$(sText, lBreak)
                ' now return and exit
                SplitWords = sWords
                Exit Function
            Else
                sWords(UBound(sWords)) = Mid$(sText, lPos, lBreak - lPos)
            End If
        End If
        ReDim Preserve sWords(UBound(sWords) + 1)
    
        ' move the pointer on a bit
        lPos = lBreak + 1
        
        ' setup the exit condition
        If lPos >= Len(sText) Then Exit Do
    
    Loop

    ' return the array
    SplitWords = sWords

End Function

'//--[RemoveEOL]------------------------------------------------------------//
'
'  Removes leading and trailing vbCrLf from strings
'
Private Function RemoveEOL(ByVal sText As String) As String
Dim sTmp As String
    ' remove leading or trailing vbCrLf from the string
    sTmp = sText
    If Left$(sTmp, 2) = vbCrLf Then
        sTmp = Right$(sTmp, Len(sTmp) - 2)
    End If
    If Right$(sTmp, 2) = vbCrLf Then
        sTmp = Left$(sTmp, Len(sTmp) - 2)
    End If
    RemoveEOL = sTmp
End Function

'//--[RemoveStrings]-------------------------------------------------------//
'
'  Removes any quoted strings from the text, but only
'  those that aren't within comments of course.
'
Private Function RemoveStrings(ByVal sText As String) As String
Dim lCom As Long
Dim lPos As Long
Dim lPos2 As Long

    lCom = InStr(1, sText, CHAR_COMMENT)
    lPos = InStr(1, sText, Chr$(34))
    If lPos < lCom Or lCom = 0 Then
        Do While lPos <> 0
            ' find the end " char to make a pair
            lPos2 = InStr(lPos + 1, sText, Chr$(34))
            If lPos2 <> 0 Then
                ' we've found a pair, so remove it
                sText = Mid$(sText, 1, lPos - 1) & Mid$(sText, lPos2 + 1)
                ' find the next starting " avoiding
                ' comments within strings
                lCom = InStr(lPos2 + 1, sText, CHAR_COMMENT)
                lPos = InStr(lPos2 + 1, sText, Chr$(34))
                If lPos > lCom Then Exit Do
            Else
                Exit Do
            End If
        Loop
    End If
    
    ' return
    RemoveStrings = sText
    
End Function

'//--[CombSort]------------------------------------------------------------//
'
'  This is a standard comb sort - you could replace
'  this with any other sorting algorithm, I just prefer
'  this one because a) i wrote it :), and b) it performs
'  well across all ranges of input arrays - it makes
'  no assumptions about how sorted the array already
'  is, because it doesn't matter.
'  The comb sort is a slight variation on the bubblesort,
'  and i know what you're thinking - ewwww, bubble sorts -
'  but you'd be wrong, the comb is only fractionally
'  slower than a quicksort... so enjoy!
'  for more on the combsort, read here:
'  http://yagni.com/combsort/index.php
'  http://cs.clackamas.cc.or.us/molatore/cs260Spr01/combsort.htm
'
Private Sub CombSort(Arr() As WORD_TYPE)
Dim i As Long, j As Long, t As WORD_TYPE
Dim swapped As Boolean
Dim gap As Long
   
    gap = UBound(Arr)
    
    Do
        gap = (gap * 10) \ 13
        If gap = 9 Or gap = 10 Then gap = 11
        If gap < 1 Then gap = 1
        
        swapped = False
        For i = 0 To UBound(Arr) - gap
            j = i + gap
            If Arr(i).Text > Arr(j).Text Then
                LSet t = Arr(j)
                LSet Arr(j) = Arr(i)
                LSet Arr(i) = t
                swapped = True
            End If
        Next i
        
        If (gap = 1) And (Not swapped) Then Exit Do
    Loop
    
End Sub


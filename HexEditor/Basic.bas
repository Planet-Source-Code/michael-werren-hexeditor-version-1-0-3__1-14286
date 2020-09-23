Attribute VB_Name = "Basic"
' Name              : HexEditor
' Author            : Michael Werren, mike@werren.com
' Source            : Visual Basic 6 with SP4
' Start             : 21.12.2000
' Last Modification : 9.1.2001

Option Explicit
Public Type FileInfo
  Name As String
  Size As Long
End Type


Public giHexSelPointer As Integer   ' Pointer to the Line in Listbox
Public gbChangedFile As Boolean     ' Pointer to Change Code of File
Public gbChangedField As Boolean    ' Pointer to change code of Fields
Public giOpt00B7 As Integer         ' Pointer to the 00 or B7 Option
Public gsOldValue As String         ' Store the old value of a Hex or Ascii field
Public gtFileInfo As FileInfo       ' Store the information of open file
Public glLastFound As Long
Sub FindNext()
  Dim lsSearchText As String
  Dim liX As Integer
  Dim lbHexSearch As Boolean
  Dim lbMatchCase As Boolean
  
  Const cB7Value = 183
  Const cFindLine = 3
  
  If frmFind.txtSearch.Text <> "" Then
    If frmFind.optFillB7.Value = vbChecked Then
      lsSearchText = ""
      For liX = 1 To Len(frmFind.txtSearch.Text)
        lsSearchText = lsSearchText + Mid(frmFind.txtSearch.Text, liX, 1) + Chr(cB7Value)
      Next liX
      ' Cut the lase B7 Value
      lsSearchText = Left(lsSearchText, Len(lsSearchText) - 1)
    Else
      lsSearchText = frmFind.txtSearch.Text
    End If
    
    ' Hex or ASCII Server
    If frmFind.optSearchIn(0).Value = True Then
      lbHexSearch = False
    Else
      lbHexSearch = True
    End If
    
    If frmFind.optMatchCase.Value = vbChecked Then
      lbMatchCase = True
    Else
      lbMatchCase = False
    End If
    
    ' Set the Start
    If frmMain.LstHexView.ListItems.Count >= frmMain.LstHexView.SelectedItem.Index + 1 Then
      glLastFound = frmMain.LstHexView.SelectedItem.Index + 1
      
      ' Start the search
      If frmFind.optSearchTyp(0).Value = True Then
        ' Simple Search
        glLastFound = SearchText(glLastFound, lsSearchText, lbHexSearch, lbMatchCase, 1)
        If glLastFound = -1 Then
          ' Notfound
          glLastFound = frmMain.LstHexView.SelectedItem.Index
        Else
          ' Found it
          With frmMain.LstHexView.ListItems.Item(glLastFound)
            .EnsureVisible
            .Selected = True
          End With
          glLastFound = glLastFound + 1
        End If
      Else
        ' Extended Search
        glLastFound = SearchText(glLastFound, lsSearchText, lbHexSearch, lbMatchCase, cFindLine)
        If glLastFound = -1 Then
          ' Notfound
          glLastFound = frmMain.LstHexView.SelectedItem.Index
        Else
          ' Found it
          With frmMain.LstHexView.ListItems.Item(glLastFound)
            .EnsureVisible
            .Selected = True
          End With
          glLastFound = glLastFound + 1
          frmFind.Hide
        End If
      End If
    End If
  End If

End Sub

Function SearchIndex(SearchString) As Long
  Dim itmfound As ListItem
  Set itmfound = frmMain.LstHexView.FindItem(SearchString, lvwText, , lvwPartial)

  itmfound.EnsureVisible          ' Scroll ListView to show found ListItem.
  itmfound.Selected = True        ' Select the ListItem.
  SearchIndex = frmMain.LstHexView.SelectedItem.Index
  Set itmfound = Nothing
End Function

Sub SetIndex(SearchString)
  Dim itmfound As ListItem
  Set itmfound = frmMain.LstHexView.FindItem(SearchString, lvwText, , lvwPartial)

  itmfound.EnsureVisible          ' Scroll ListView to show found ListItem.
  itmfound.Selected = True        ' Select the ListItem.
  frmMain.LstHexView.SetFocus     ' Set the focus to the object
  Set itmfound = Nothing
End Sub

Function SearchText(StartLine As Long, SearchString As String, HexMode As Boolean, MatchCase As Boolean, TotSearch As Integer) As Long
  Dim llLineCounter As Long
  Dim llMaxCounter As Long
  Dim lbFound As Boolean
  Dim lsOrgText As String
  Dim liMaxLine As Integer
  
  llMaxCounter = frmMain.LstHexView.ListItems.Count
  llLineCounter = StartLine
  lbFound = False
  
  While (llLineCounter <= llMaxCounter) And Not lbFound
   
    If HexMode Then
      ' Search in Hexline, Subitems(1)
      If TotSearch = 1 Then
        lsOrgText = frmMain.LstHexView.ListItems.Item(llLineCounter).SubItems(1)
      Else
        If llLineCounter + TotSearch > frmMain.LstHexView.ListItems.Count Then
          TotSearch = frmMain.LstHexView.ListItems.Count - frmMain.LstHexView.ListItems.Count - 1
        End If
        lsOrgText = ""
        For liMaxLine = llLineCounter To llLineCounter + TotSearch
          lsOrgText = lsOrgText + frmMain.LstHexView.ListItems.Item(liMaxLine).SubItems(1) + "."
        Next liMaxLine
      End If
      lsOrgText = UCase(lsOrgText)
      SearchString = UCase(SearchString)
      ' Replace the doublespace with .
      While InStr(lsOrgText, "  ") <> 0
        lsOrgText = Trim(Mid(lsOrgText, 1, InStr(lsOrgText, " ") - 1)) + "." + _
                    Trim(Mid(lsOrgText, InStr(lsOrgText, " "), Len(lsOrgText)))
      Wend
      If InStr(lsOrgText, SearchString) <> 0 Then
        ' Found it
        SearchText = llLineCounter
        ' Found it in expanded mode
        If TotSearch <> 1 Then
          SearchText = llLineCounter + InStr(lsOrgText, SearchString) \ 48
        End If
        lbFound = True
      End If
    Else
      ' Search in ASCII, Subitems(2)
      If TotSearch = 1 Then
        lsOrgText = frmMain.LstHexView.ListItems.Item(llLineCounter).SubItems(2)
      Else
        If llLineCounter + TotSearch > frmMain.LstHexView.ListItems.Count Then
          TotSearch = frmMain.LstHexView.ListItems.Count - frmMain.LstHexView.ListItems.Count - 1
        End If
        lsOrgText = ""
        For liMaxLine = llLineCounter To llLineCounter + TotSearch
          lsOrgText = lsOrgText + frmMain.LstHexView.ListItems.Item(liMaxLine).SubItems(2)
        Next liMaxLine
      End If
      If MatchCase Then
        lsOrgText = UCase(lsOrgText)
        SearchString = UCase(SearchString)
      End If
      If InStr(lsOrgText, SearchString) <> 0 Then
        ' Found it
        SearchText = llLineCounter
        ' Found it in expanded mode
        If TotSearch <> 1 Then
          SearchText = llLineCounter + InStr(lsOrgText, SearchString) \ 16
        End If
        lbFound = True
      End If
    End If
    llLineCounter = llLineCounter + 1
  Wend
 
  If Not lbFound Then
    MsgBox "Can not find the String", vbInformation + vbOKOnly, "Search for: " + SearchString
    SearchText = -1
  End If
End Function

Sub DispMsg(Text As String)
  frmMain.StatusBar.Panels(1).Text = Text
End Sub

Function SpaceFill(Counter As Integer, Char As String) As String
  SpaceFill = String(Counter, Char)
End Function

Sub AddHexLine(HexIndex As String, HexText As String, AsciiText As String)
  Dim itmX As ListItem
  
  Set itmX = frmMain.LstHexView.ListItems.Add
  itmX.Text = HexIndex
  itmX.SubItems(1) = HexText
  itmX.SubItems(2) = AsciiText
End Sub

Function SaveTranslate() As String
  Dim lsTextLine As String
  Dim lsHexLine As String
  Dim lsHexCode As String
  Dim lsTextSave As String
  Dim liX As Integer
  Dim llX As Long
  Dim llY As Long
  Dim llMax As Long
  Dim liprocentold As Integer
  Dim liprocent As Integer
  
  Const cB7Value = 183
  
  Screen.MousePointer = vbHourglass
  llMax = frmMain.LstHexView.ListItems.Count
  lsTextSave = String(gtFileInfo.Size, Chr(0))
  llY = 1
  
  For llX = 1 To llMax
    ' Replace the doublespace with .
    lsTextLine = frmMain.LstHexView.ListItems.Item(llX).SubItems(1)
    If InStr(lsTextLine, "  ") <> 0 Then
      lsTextLine = Trim(Mid(lsTextLine, 1, InStr(lsTextLine, " ") - 1)) + "." + _
                   Trim(Mid(lsTextLine, InStr(lsTextLine, " "), Len(lsTextLine)))
    End If
    
    lsHexLine = ""
    While InStr(lsTextLine, ".") <> 0
      lsHexCode = Mid(lsTextLine, 1, InStr(lsTextLine, ".") - 1)
      lsHexLine = lsHexLine + Chr(Hex2Dec(lsHexCode))
      lsTextLine = Mid(lsTextLine, InStr(lsTextLine, ".") + 1, Len(lsTextLine))
    Wend
    ' do not forget the last byte
    lsHexCode = Trim(lsTextLine)
    lsHexLine = lsHexLine + Chr(Hex2Dec(lsHexCode))
        
    For liX = 1 To Len(lsHexLine)
      Mid(lsTextSave, llY, 1) = Mid(lsHexLine, liX, 1)
      llY = llY + 1
    Next liX
    
    liprocentold = liprocent
    liprocent = llX * 100 \ llMax
    If liprocent <> liprocentold Then
      frmMain.ProBar.Value = liprocent
    End If
  Next llX
  
  SaveTranslate = lsTextSave
  frmMain.ProBar.Value = 0
  Screen.MousePointer = vbDefault
End Function

Sub HexTranslate(TransText As String)
  Dim lsTransText As String
  Dim lsZeichen As String
  Dim lsOrgText As String
  Dim lsHexCode As String
  Dim lsHexLine As String
  Dim lsHexIndex As String
  Dim liHexIndex As Long
  Dim liZeichen As Integer
  Dim liPointer As Integer
  Dim liX As Long
  Dim liprocent As Integer
  Dim liprocentold As Integer

  lsTransText = TransText
  
  ' Start Translation
  Screen.MousePointer = vbHourglass
  DispMsg "Translate the file, please wait..."
  liPointer = 1
  liHexIndex = 0

  For liX = 1 To Len(lsTransText)
    If liPointer <= 16 Then
      liPointer = liPointer + 1
      lsZeichen = Mid(lsTransText, liX, 1)
      liZeichen = Asc(lsZeichen)
      lsHexCode = Hex(liZeichen)
      
      If Len(lsHexCode) < 2 Then
        lsHexCode = "0" + lsHexCode
      End If
      If liPointer <= 16 Then
        If liPointer <> 9 Then
          lsHexLine = lsHexLine + lsHexCode + "."
        Else
          lsHexLine = lsHexLine + lsHexCode + "  "
        End If
      Else
        lsHexLine = lsHexLine + lsHexCode
        ' Enum the translation in procent
        liprocentold = liprocent
        liprocent = liX * 100 \ Len(lsTransText)
        If liprocent <> liprocentold Then
          frmMain.ProBar.Value = liprocent
        End If
      End If
      If Asc(lsZeichen) = 0 Then
        lsOrgText = lsOrgText + "·"
      Else
        lsOrgText = lsOrgText + lsZeichen
      End If
      
    Else
      lsHexIndex = SpaceFill(8 - Len(Hex(liHexIndex)), "0") + Hex(liHexIndex)
      AddHexLine lsHexIndex, lsHexLine, lsOrgText
      liPointer = 1
      liHexIndex = liHexIndex + 16
      lsHexLine = ""
      lsOrgText = ""
      liX = liX - 1
    End If
  Next liX
  
  ' Is ther still a recorde to add ?
  If lsHexLine <> "" Then
    If Mid(lsHexLine, Len(lsHexLine), 1) = "." Then
      lsHexLine = Mid(lsHexLine, 1, Len(lsHexLine) - 1)
    End If
    lsHexIndex = SpaceFill(8 - Len(Hex(liHexIndex)), "0") + Hex(liHexIndex)
    AddHexLine lsHexIndex, lsHexLine, lsOrgText
  End If
  DispMsg "View mode"
  frmMain.ProBar.Value = 0
  Screen.MousePointer = vbDefault
End Sub

Function Hex2Dec(HexValue As String) As Integer
  Dim lsChar As String
  Dim liHBit As Integer
  Dim liLBit As Integer
  
  ' H Byte
  lsChar = Mid(HexValue, 1, 1)
  Select Case Asc(lsChar)
    Case 48 To 57 '0..9
      liHBit = Val(lsChar) * 16
    Case 65 To 70 'A..F
      liHBit = (((65 - Asc(lsChar)) * -1) + 10) * 16
  End Select
    
  ' L Byte
  lsChar = Mid(HexValue, 2, 1)
  Select Case Asc(lsChar)
    Case 48 To 57 '0..9
      liLBit = Val(lsChar)
    Case 65 To 70 'A..F
      liLBit = (65 - Asc(lsChar)) * -1 + 10
  End Select
  
  Hex2Dec = liHBit + liLBit
End Function

Sub LoadHexBlock(LineIndex As Integer, MaxLineIndex As Integer)
  Dim liX As Integer
  Dim liY As Integer
  Dim lsText As String
  Dim liScrollValue As Integer
  Dim liMaxRecord As Integer
  Dim liHexSourcePointer As Integer
  Dim lsHexIndex As String
  Dim lsHexCode As String
  Dim lsAsciiCode As String
    
  Const cMaxRecord = 19 ' 0-19 = 20 Records
  
  ' Reset the HexFields
  For liX = 0 To 319
    With frmMain.txtHex(liX)
      .Text = ""
      .Enabled = False
      .BackColor = &H8000000F
    End With
  Next liX
  
  ' Reset the AsciiFields
  For liX = 0 To cMaxRecord
    With frmMain.txtAscii(liX)
      .Text = ""
      .Enabled = False
      .BackColor = &H8000000F
    End With
  Next liX
  
  ' Load Data records
  If LineIndex < 1 Then LineIndex = 1
  giHexSelPointer = LineIndex
  liScrollValue = LineIndex
  If LineIndex <= cMaxRecord Then
    LineIndex = 1
    liScrollValue = 1
  End If
  
  ' Are the other Hex records ?
    If MaxLineIndex >= giHexSelPointer + cMaxRecord Then
      liMaxRecord = cMaxRecord
    Else
      If MaxLineIndex <= cMaxRecord Then
        giHexSelPointer = 1
        liMaxRecord = MaxLineIndex - 1
      Else
        giHexSelPointer = MaxLineIndex - cMaxRecord
        liMaxRecord = cMaxRecord
      End If
    End If
  
  For liX = 0 To liMaxRecord
    liHexSourcePointer = giHexSelPointer + liX
    lsHexIndex = frmMain.LstHexView.ListItems.Item(liHexSourcePointer).Text
    lsHexCode = frmMain.LstHexView.ListItems.Item(liHexSourcePointer).SubItems(1)
    lsAsciiCode = frmMain.LstHexView.ListItems.Item(liHexSourcePointer).SubItems(2)
    
    ' Replace the doublespace with .
    If InStr(lsHexCode, "  ") <> 0 Then
      lsHexCode = Trim(Mid(lsHexCode, 1, InStr(lsHexCode, " ") - 1)) + "." + _
                  Trim(Mid(lsHexCode, InStr(lsHexCode, " "), Len(lsHexCode)))
    End If
    ' Split the Hexcode an add to the mask
    liY = -1
    Do
      liY = liY + 1
      If InStr(lsHexCode, ".") <> 0 Then
        lsText = Trim(Mid(lsHexCode, 1, InStr(lsHexCode, ".") - 1))
        lsHexCode = Trim(Mid(lsHexCode, InStr(lsHexCode, ".") + 1, Len(lsHexCode)))
      Else
        lsText = Trim(lsHexCode)
        lsHexCode = ""
      End If
      frmMain.txtHex(liX * 16 + liY).Text = lsText
      frmMain.txtHex(liX * 16 + liY).ForeColor = &H80000008
      frmMain.txtHex(liX * 16 + liY).BackColor = &H80000005
      frmMain.txtHex(liX * 16 + liY).Enabled = True
    Loop Until Len(lsHexCode) = 0
    
    ' Add the Index
    With frmMain.txtHexIndex(liX)
      .Text = lsHexIndex
      .ForeColor = &H80000008
    End With
    ' Add the Ascii
    If Len(lsAsciiCode) < 16 Then
      lsAsciiCode = lsAsciiCode + SpaceFill(16 - Len(lsAsciiCode), " ")
    End If
    With frmMain.txtAscii(liX)
      .Text = lsAsciiCode
      .ForeColor = &H80000008
      .BackColor = &H80000005
      .Enabled = True
    End With
  Next liX
  
  ' Show mode
  DispMsg "Editor mode"
  ' initialized the VScrollBar max
  liX = frmMain.LstHexView.ListItems.Count - cMaxRecord
  If liX < 1 Then liX = 1
  frmMain.VScrollEditor.Max = liX
  frmMain.VScrollEditor.Value = giHexSelPointer
  ' set focus to  the first object
  If frmMain.FrameEditor.Visible Then
    frmMain.txtHex(0).SetFocus
  End If
End Sub

Sub SetMenu(MenuMode As String)
  If UCase(MenuMode) = "VIEWER" Then
    frmMain.mnuFileLoad.Enabled = True
    frmMain.Toolbar1.Buttons(1).Enabled = True
    frmMain.Toolbar1.Buttons(1).ToolTipText = "Load a file"
    frmMain.mnuFileSave.Enabled = True
    frmMain.Toolbar1.Buttons(2).Enabled = True
    frmMain.Toolbar1.Buttons(2).ToolTipText = "Save a file"
    frmMain.mnuFileOptions.Enabled = True
    frmMain.mnuFileInfo.Enabled = True
    frmMain.mnuFileExit.Enabled = True
    frmMain.Toolbar1.Buttons(12).Enabled = True
    frmMain.Toolbar1.Buttons(12).ToolTipText = "Exit the program"
    
    frmMain.mnuViewerFind.Enabled = True
    frmMain.Toolbar1.Buttons(8).Enabled = True
    frmMain.Toolbar1.Buttons(8).ToolTipText = "Find..."
    If frmFind.txtSearch.Text <> "" Then
      frmMain.mnuViewerNext.Enabled = True
    Else
      frmMain.mnuViewerNext.Enabled = False
    End If
    
    frmMain.mnuEditorEdit.Enabled = True
    frmMain.Toolbar1.Buttons(5).ToolTipText = "Editor mode"
    frmMain.mnuEditorReload.Enabled = False
    frmMain.Toolbar1.Buttons(6).Enabled = False
    frmMain.Toolbar1.Buttons(6).ToolTipText = "This option is turned off in this mode"
    
  ElseIf UCase(MenuMode) = "EDITOR" Then
    frmMain.mnuFileLoad.Enabled = False
    frmMain.Toolbar1.Buttons(1).Enabled = False
    frmMain.Toolbar1.Buttons(1).ToolTipText = "This option is turned off in this mode"
    frmMain.mnuFileSave.Enabled = False
    frmMain.Toolbar1.Buttons(2).Enabled = False
    frmMain.Toolbar1.Buttons(2).ToolTipText = "This option is turned off in this mode"
    frmMain.mnuFileOptions.Enabled = True
    frmMain.mnuFileInfo.Enabled = True
    frmMain.mnuFileExit.Enabled = False
    frmMain.Toolbar1.Buttons(12).Enabled = False
    frmMain.Toolbar1.Buttons(12).ToolTipText = "This option is turned off in this mode"

    frmMain.mnuViewerFind.Enabled = False
    frmMain.Toolbar1.Buttons(8).Enabled = False
    frmMain.Toolbar1.Buttons(8).ToolTipText = "This option is turned off in this mode"
    frmMain.mnuViewerNext.Enabled = False

    frmMain.mnuEditorEdit.Enabled = True
    frmMain.Toolbar1.Buttons(5).ToolTipText = "Viewer mode"
    frmMain.mnuEditorReload.Enabled = True
    frmMain.Toolbar1.Buttons(6).Enabled = True
    frmMain.Toolbar1.Buttons(6).ToolTipText = "Reload the Frame"
  End If
End Sub

Sub SaveBlock(StartPointer As Long)
  Dim llPosition As Long
  Dim liX As Integer
  Dim liY As Integer
  Dim liPosPointer As Integer
  Dim lsHexIndex As String
  Dim lsHex As String
  Dim lsAscii As String
  
  Const cMaxRecord = 19 ' 0-19 = 20 Records
  ' Set the Blockstate
  frmMain.StatusBar.Panels(3).Picture = frmMain.ImgRed.Picture
  
  For liX = 0 To cMaxRecord
    llPosition = StartPointer + liX
    ' Get the Index
    lsHexIndex = frmMain.txtHexIndex(liX).Text
    ' Get the Hex Values and build the String
    lsHex = ""
    For liY = 0 To 15
      liPosPointer = liX * 16 + liY
      If frmMain.txtHex(liPosPointer).Enabled Then
        If liY = 7 Then
          lsHex = lsHex + frmMain.txtHex(liPosPointer).Text + "  "
        Else
          If liY <> 15 Then
            lsHex = lsHex + frmMain.txtHex(liPosPointer).Text + "."
          Else
            lsHex = lsHex + frmMain.txtHex(liPosPointer).Text
          End If
        End If
      End If
    Next liY
    ' Get the Ascii
    lsAscii = frmMain.txtAscii(liX).Text
    If frmMain.txtAscii(liX).Enabled Then
      ' Save the Hex line
      If Mid(lsHex, Len(lsHex), 1) <> "." Then
        frmMain.LstHexView.ListItems(llPosition).SubItems(1) = lsHex
      Else
        frmMain.LstHexView.ListItems(llPosition).SubItems(1) = Left(lsHex, Len(lsHex) - 1)
      End If
      ' Save the Ascii line
      frmMain.LstHexView.ListItems(llPosition).SubItems(2) = lsAscii
    End If
  Next liX
End Sub

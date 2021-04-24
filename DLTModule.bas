Attribute VB_Name = "DLTModule"

Public Type DLTControls
    ControlHeightActive As Integer
    ControlHeightInactive As Integer
    ControlColorActive As OLE_COLOR
    ControlColorInactive As OLE_COLOR
    ControlColorDisabled As OLE_COLOR
    ControlColorCaption As OLE_COLOR
    ControlColorValid As OLE_COLOR
    ControlColorInvalid As OLE_COLOR
    ControlColorBack As OLE_COLOR
    ControlTotalTabs As Integer
    ControlTabCaptions() As String
    ControlTabEnableds() As Boolean
    ControlTabVisibles() As Boolean
    ControlTabGoodStates() As Boolean
    ControlTabWidths() As Integer
    ControlTabLefts() As Integer
    ControlTabRows() As Integer
    ControlTabControls() As String
    ControlTabsPerRow As Integer
    ControlTotalRows As Integer
    ControlRowNumbers() As Integer
    ControlActiveTab As Integer
    ControlCaptionOffset As Integer
    ControlStateMode As Integer
    ControlFlatMode As Integer
    ControlInsideBorder As Boolean
    ControlStateWidth As Integer
    ControlReadData As String
    ControlOffsetMode As Integer
    ControlOffsetWidth As Integer
End Type

Const DLTCaptionSpacing = 120

Public Sub ReadDLTPropertyBag(PropControls As DLTControls, PropBag As PropertyBag)

Dim ReadTabkey As String
Dim ReadTabnumber As Integer
Dim ReadObjectname As String
Dim ReadObjectindex As String
Dim ReadControlkey As String
Dim ReadControlnumber As Integer
Dim ReadTotalcontrols As Integer
Dim ReadName As String

Debug.Print

PropControls.ControlColorActive = PropBag.ReadProperty("ActiveColor", &H8000000F)
PropControls.ControlColorInactive = PropBag.ReadProperty("InactiveColor", &H8000000F)
PropControls.ControlColorDisabled = PropBag.ReadProperty("DisabledColor", &H8000000F)
PropControls.ControlColorBack = PropBag.ReadProperty("BackgroundColor", &H8000000F)
PropControls.ControlColorCaption = PropBag.ReadProperty("CaptionColor", &H646464)
PropControls.ControlColorInvalid = PropBag.ReadProperty("InvalidColor", &HC0C0FF)
PropControls.ControlColorValid = PropBag.ReadProperty("ValidColor", &HC0FFC0)

PropControls.ControlHeightActive = PropBag.ReadProperty("ActiveHeight", 22)
PropControls.ControlHeightInactive = PropBag.ReadProperty("InactiveHeight", 19)

PropControls.ControlStateMode = PropBag.ReadProperty("StateMode", 2)
PropControls.ControlStateWidth = PropBag.ReadProperty("StateWidth", 60)

PropControls.ControlOffsetMode = PropBag.ReadProperty("OffsetMode", 0)
PropControls.ControlOffsetWidth = PropBag.ReadProperty("OffsetWidth", 30)

PropControls.ControlFlatMode = PropBag.ReadProperty("FlatMode", 3)

PropControls.ControlInsideBorder = PropBag.ReadProperty("InsideBorder", True)

PropControls.ControlTotalTabs = PropBag.ReadProperty("TotalTabs", 1)
PropControls.ControlTabsPerRow = PropBag.ReadProperty("TabsPerRow", 0)
PropControls.ControlActiveTab = PropBag.ReadProperty("ActiveTab", 0)

If PropControls.ControlTotalTabs < 1 Then PropControls.ControlTotalTabs = 1
If PropControls.ControlTabsPerRow < 0 Then PropControls.ControlTabsPerRow = 0
If PropControls.ControlActiveTab < 0 Or PropControls.ControlActiveTab >= PropControls.ControlTotalTabs Then PropControls.ControlActiveTab = 0

Debug.Print "READ= TotalTabs " & PropControls.ControlTotalTabs
Debug.Print "READ= TabsPerRow " & PropControls.ControlTabsPerRow
Debug.Print "READ= ActiveTab " & PropControls.ControlActiveTab

Call RedimTabArrays(PropControls)

PropControls.ControlReadData = ""

ReadTabnumber = 0
While ReadTabnumber < PropControls.ControlTotalTabs

    ReadTabkey = "Tab(" + Trim$(Str$(ReadTabnumber)) + ")"

    ReadName = ReadTabkey + ".Caption"
    PropControls.ControlTabCaptions(ReadTabnumber) = PropBag.ReadProperty(ReadName, "")
    Debug.Print "READ= " & ReadName & " " & PropControls.ControlTabCaptions(ReadTabnumber)

    ReadName = ReadTabkey + ".Enabled"
    PropControls.ControlTabEnableds(ReadTabnumber) = PropBag.ReadProperty(ReadName, True)
    Debug.Print "READ= " & ReadName & " " & PropControls.ControlTabEnableds(ReadTabnumber)

    ReadName = ReadTabkey + ".Visible"
    PropControls.ControlTabVisibles(ReadTabnumber) = PropBag.ReadProperty(ReadName, True)
    Debug.Print "READ= " & ReadName & " " & PropControls.ControlTabVisibles(ReadTabnumber)

    ReadName = ReadTabkey + ".GoodState"
    PropControls.ControlTabGoodStates(ReadTabnumber) = PropBag.ReadProperty(ReadName, True)
    Debug.Print "READ= " & ReadName & " " & PropControls.ControlTabGoodStates(ReadTabnumber)
    
    PropControls.ControlTabControls(ReadTabnumber) = ""

    ReadName = ReadTabkey + ".TotalSubcontrols"
    ReadTotalcontrols = PropBag.ReadProperty(ReadName, True)
    Debug.Print "READ= " & ReadName & " " & ReadTotalcontrols

    ReadControlnumber = 0
    While ReadControlnumber < ReadTotalcontrols
    
        ReadControlkey = ReadTabkey + ".Subcontrol(" + Trim$(Str$(ReadControlnumber)) + ")"

        ReadName = ReadControlkey + ".Name"
        ReadObjectname = PropBag.ReadProperty(ReadName, "")
        Debug.Print "READ= " & ReadName & " " & ReadObjectname

        ReadName = ReadControlkey + ".Index"
        ReadObjectindex = PropBag.ReadProperty(ReadName, "")
        Debug.Print "READ= " & ReadName & " " & ReadObjectindex

        PropControls.ControlReadData = PropControls.ControlReadData + Trim$(Str$(ReadTabnumber)) + "|" + ReadObjectname + "|" + ReadObjectindex + "|"

        ReadControlnumber = ReadControlnumber + 1
    Wend
    
    ReadTabnumber = ReadTabnumber + 1
Wend

Debug.Print "DATA= " & PropControls.ControlReadData

End Sub

Public Sub WriteDLTPropertyBag(PropMe As DataLTab, PropControls As DLTControls, PropBag As PropertyBag)

Dim WriteTabkey As String
Dim WriteTabnumber As Integer
Dim WriteControlnumber As Integer
Dim WriteControlhandles As String
Dim WriteControldata As String
Dim WriteControlkey As String
Dim WriteItemtotal As Integer
Dim WriteItemindex As Integer
Dim WriteObject As Control
Dim WriteHandle As String
Dim WriteName As String
Dim WriteValue As Long

Debug.Print

Call PropBag.WriteProperty("ActiveColor", PropControls.ControlColorActive, &H8000000F)
Call PropBag.WriteProperty("InactiveColor", PropControls.ControlColorInactive, &H8000000F)
Call PropBag.WriteProperty("DisabledColor", PropControls.ControlColorDisabled, &H8000000F)
Call PropBag.WriteProperty("BackgroundColor", PropControls.ControlColorBack, &H8000000F)
Call PropBag.WriteProperty("CaptionColor", PropControls.ControlColorCaption, &H646464)
Call PropBag.WriteProperty("InvalidColor", PropControls.ControlColorInvalid, &HC0C0FF)
Call PropBag.WriteProperty("ValidColor", PropControls.ControlColorValid, &HC0FFC0)

Call PropBag.WriteProperty("ActiveHeight", PropControls.ControlHeightActive, 22)
Call PropBag.WriteProperty("InactiveHeight", PropControls.ControlHeightInactive, 19)

Call PropBag.WriteProperty("StateMode", PropControls.ControlStateMode, 2)
Call PropBag.WriteProperty("StateWidth", PropControls.ControlStateWidth, 60)

Call PropBag.WriteProperty("OffsetMode", PropControls.ControlOffsetMode, 0)
Call PropBag.WriteProperty("OffsetWidth", PropControls.ControlOffsetWidth, 30)

Call PropBag.WriteProperty("FlatMode", PropControls.ControlFlatMode, 3)

Call PropBag.WriteProperty("InsideBorder", PropControls.ControlInsideBorder, True)

Call PropBag.WriteProperty("TotalTabs", PropControls.ControlTotalTabs, 1)
Debug.Print "WRITE= TotalTabs " & PropControls.ControlTotalTabs
Call PropBag.WriteProperty("TabsPerRow", PropControls.ControlTabsPerRow, 0)
Debug.Print "WRITE= TabsPerRow " & PropControls.ControlTabsPerRow
Call PropBag.WriteProperty("ActiveTab", PropControls.ControlActiveTab, 0)
Debug.Print "WRITE= ActiveTabs " & PropControls.ControlActiveTab

PropMe.ContainedControlsCount WriteItemtotal

WriteTabnumber = 0
While WriteTabnumber < PropControls.ControlTotalTabs

    WriteTabkey = "Tab(" + Trim$(Str$(WriteTabnumber)) + ")"
    
    WriteName = WriteTabkey + ".Caption"
    Call PropBag.WriteProperty(WriteName, PropControls.ControlTabCaptions(WriteTabnumber), "")
    Debug.Print "WRITE= " & WriteName & " " & PropControls.ControlTabCaptions(WriteTabnumber)

    WriteName = WriteTabkey + ".Enabled"
    Call PropBag.WriteProperty(WriteName, PropControls.ControlTabEnableds(WriteTabnumber), True)
    Debug.Print "WRITE= " & WriteName & " " & PropControls.ControlTabEnableds(WriteTabnumber)

    WriteName = WriteTabkey + ".Visible"
    Call PropBag.WriteProperty(WriteName, PropControls.ControlTabVisibles(WriteTabnumber), True)
    Debug.Print "WRITE= " & WriteName & " " & PropControls.ControlTabVisibles(WriteTabnumber)

    WriteName = WriteTabkey + ".GoodState"
    Call PropBag.WriteProperty(WriteName, PropControls.ControlTabGoodStates(WriteTabnumber), True)
    Debug.Print "WRITE= " & WriteName & " " & PropControls.ControlTabGoodStates(WriteTabnumber)
    
    WriteControlhandles = "|" + PropControls.ControlTabControls(WriteTabnumber)
    WriteControlnumber = 0
    
    WriteItemindex = 0
    While WriteItemindex < WriteItemtotal
    
        PropMe.ContainedControlsItem WriteItemindex, WriteObject

        On Error Resume Next
        WriteValue = -1
        WriteValue = WriteObject.hWnd
        On Error GoTo 0
        
        If WriteValue > 0 Then
        
            WriteHandle = "|" + Trim$(Str$(WriteValue)) + "|"

            If InStr(WriteControlhandles, WriteHandle) > 0 Then
        
                WriteControlkey = WriteTabkey + ".Subcontrol(" + Trim$(Str$(WriteControlnumber)) + ")"
            
                WriteControldata = WriteObject.Name
                WriteName = WriteControlkey + ".Name"
                Call PropBag.WriteProperty(WriteName, WriteControldata, "")
                Debug.Print "WRITE= " & WriteName & " " & WriteControldata
            
                WriteControldata = GetObjectIndex(WriteObject)
                WriteName = WriteControlkey + ".Index"
                Call PropBag.WriteProperty(WriteName, WriteControldata, "")
                Debug.Print "WRITE= " & WriteName & " " & WriteControldata
            
                WriteControlnumber = WriteControlnumber + 1
            
            End If
        
        End If
        
        WriteItemindex = WriteItemindex + 1
    Wend

    WriteName = WriteTabkey + ".TotalSubcontrols"
    Call PropBag.WriteProperty(WriteName, WriteControlnumber, 0)
    Debug.Print "WRITE= " & WriteName & " " & WriteControlnumber

    WriteTabnumber = WriteTabnumber + 1
Wend

End Sub

Public Sub InitializeDLTControls(InitializeControls As DLTControls)

InitializeControls.ControlHeightActive = 330
InitializeControls.ControlHeightInactive = 285

InitializeControls.ControlStateMode = 2
InitializeControls.ControlStateWidth = 60

InitializeControls.ControlOffsetMode = 0
InitializeControls.ControlOffsetWidth = 30

InitializeControls.ControlFlatMode = 3

InitializeControls.ControlInsideBorder = True

InitializeControls.ControlColorActive = &H8000000F
InitializeControls.ControlColorInactive = &H8000000F
InitializeControls.ControlColorDisabled = &H8000000F
InitializeControls.ControlColorBack = &H8000000F
InitializeControls.ControlColorCaption = &H646464
InitializeControls.ControlColorValid = &HC0FFC0
InitializeControls.ControlColorInvalid = &HC0C0FF

InitializeControls.ControlTotalTabs = 1
InitializeControls.ControlTabsPerRow = 0
InitializeControls.ControlActiveTab = 0

Call RedimTabArrays(InitializeControls)

InitializeControls.ControlTabCaptions(0) = ""
InitializeControls.ControlTabEnableds(0) = True
InitializeControls.ControlTabVisibles(0) = True
InitializeControls.ControlTabGoodStates(0) = True
InitializeControls.ControlTabControls(0) = ""

End Sub

Public Sub InitPropertiesDLTControls(InitControls As DLTControls)

InitControls.ControlTotalTabs = 1
InitControls.ControlActiveTab = 0

Call RedimTabArrays(InitControls)

InitControls.ControlTabCaptions(0) = ""
InitControls.ControlTabEnableds(0) = True
InitControls.ControlTabVisibles(0) = True
InitControls.ControlTabGoodStates(0) = True
InitControls.ControlTabControls(0) = ""

End Sub

Public Sub DoMouseMove(MouseMe As DataLTab, MouseControls As DLTControls, MouseButton As Integer, MouseShift As Integer, MouseX As Single, MouseY As Single)

Dim MouseTab As Integer
Dim MouseIndex As Integer
Dim MouseInactive As Integer
Dim MouseLimit As Integer
Dim MouseLast As Integer
Dim MouseRow As Integer

MouseLast = MouseControls.ControlTotalRows - 1
MouseLimit = MouseLast * MouseControls.ControlHeightInactive + MouseControls.ControlHeightActive

If MouseY > MouseLimit Then Exit Sub

MouseInactive = MouseLimit - MouseControls.ControlHeightInactive
MouseLimit = MouseLimit - MouseControls.ControlHeightActive
MouseRow = MouseControls.ControlTotalRows
MouseTab = -1

While MouseRow > 0 And MouseTab = -1
    MouseRow = MouseRow - 1
    
    MouseIndex = 0
    While MouseIndex < MouseControls.ControlTotalTabs And MouseTab = -1
    
        If MouseControls.ControlTabRows(MouseIndex) = MouseControls.ControlRowNumbers(MouseRow) Then
            If MouseX > MouseControls.ControlTabLefts(MouseIndex) And MouseX < MouseControls.ControlTabLefts(MouseIndex) + MouseControls.ControlTabWidths(MouseIndex) Then
                
                If MouseIndex = MouseControls.ControlActiveTab Or MouseRow < MouseLast Then
                    If MouseY > MouseLimit Then MouseTab = MouseIndex
                Else
                    If MouseY > MouseInactive Then MouseTab = MouseIndex
                End If
                
            End If
        End If
        
        MouseIndex = MouseIndex + 1
    Wend
    
    MouseLimit = MouseLimit - MouseControls.ControlHeightInactive
Wend

Debug.Print "MOUSEDOWN= " & MouseTab

If MouseTab < 0 Then Exit Sub

If MouseControls.ControlActiveTab = MouseTab Then Exit Sub
If MouseControls.ControlTabEnableds(MouseTab) = False Then Exit Sub

Call SetActiveTab(MouseMe, MouseControls, MouseTab)

End Sub

Public Sub DoRedrawAllTabs(RedrawMe As DataLTab, RedrawControls As DLTControls)

Dim RedrawTab As Integer
Dim RedrawWidth As Integer
Dim RedrawMaximum As Integer
Dim RedrawVisible As Boolean
Dim RedrawOffset As Integer
Dim RedrawState As Integer
Dim RedrawRow As Integer

    RedrawState = RedrawControls.ControlStateWidth
    RedrawOffset = RedrawControls.ControlOffsetWidth
    
    If RedrawState < 60 Then RedrawState = 60
    If RedrawOffset < 0 Or RedrawControls.ControlOffsetMode = 0 Then RedrawOffset = 0
    
    If RedrawControls.ControlStateMode > 1 And (RedrawControls.ControlFlatMode And 2) = 0 Then RedrawState = RedrawState + 30
    
    RedrawControls.ControlCaptionOffset = RedrawMe.TextHeight("X")
    RedrawControls.ControlCaptionOffset = RedrawControls.ControlHeightInactive - RedrawControls.ControlCaptionOffset
    RedrawControls.ControlCaptionOffset = RedrawControls.ControlCaptionOffset / 2

    RedrawMaximum = RedrawMe.ScaleWidth - 15

    If RedrawControls.ControlTabsPerRow > 0 Then

        If RedrawControls.ControlOffsetMode = 0 Then
            RedrawWidth = RedrawMaximum
        Else
            RedrawWidth = RedrawControls.ControlOffsetWidth
            If RedrawWidth < 0 Then RedrawWidth = 0
            RedrawMaximum = RedrawMaximum - RedrawWidth
            RedrawWidth = RedrawMaximum - RedrawWidth
        End If
        
        RedrawWidth = RedrawWidth / RedrawControls.ControlTabsPerRow
        RedrawVisible = False
        
    Else

        RedrawWidth = DLTCaptionSpacing * 2
        RedrawVisible = True
        
        If RedrawControls.ControlInsideBorder = True Then RedrawWidth = RedrawWidth + 30
        If (RedrawControls.ControlFlatMode And 1) = 0 Then RedrawWidth = RedrawWidth + 60
        
        Select Case RedrawControls.ControlStateMode
            Case 1:    RedrawWidth = RedrawWidth + RedrawOffset * 2
            Case 2, 3: RedrawWidth = RedrawWidth + RedrawState + RedrawOffset
        End Select

        If RedrawControls.ControlStateMode > 0 And (RedrawControls.ControlFlatMode And 2) = 0 Then RedrawWidth = RedrawWidth + 30

    End If

loopredrawcalculation:

    RedrawRow = 0

    If RedrawControls.ControlOffsetMode = 0 Then
    
        RedrawControls.ControlTabLefts(0) = 0

    Else

        RedrawControls.ControlTabLefts(0) = RedrawControls.ControlOffsetWidth

        If RedrawControls.ControlTabLefts(0) < 0 Then RedrawControls.ControlTabLefts(0) = 0

    End If
    
    For RedrawTab = 0 To RedrawControls.ControlTotalTabs - 1

        If RedrawVisible = False And RedrawControls.ControlTabVisibles(RedrawTab) = False Then
    
            RedrawControls.ControlTabWidths(RedrawTab) = 0
    
        ElseIf RedrawControls.ControlTabsPerRow > 0 Then
        
            RedrawControls.ControlTabWidths(RedrawTab) = RedrawWidth
    
        Else
        
            RedrawControls.ControlTabWidths(RedrawTab) = RedrawMe.TextWidth(ReturnTabCaption(RedrawControls, RedrawTab)) + RedrawWidth
    
            If RedrawControls.ControlTabWidths(RedrawTab) > RedrawMaximum Then RedrawControls.ControlTabWidths(RedrawTab) = RedrawMaximum
    
        End If
    
        If RedrawTab > 0 Then RedrawControls.ControlTabLefts(RedrawTab) = RedrawControls.ControlTabLefts(RedrawTab - 1) + RedrawControls.ControlTabWidths(RedrawTab - 1)
    
        If RedrawControls.ControlTabLefts(RedrawTab) + RedrawControls.ControlTabWidths(RedrawTab) > RedrawMaximum Then
        
            RedrawControls.ControlTabLefts(RedrawTab) = RedrawControls.ControlTabLefts(0)
        
            RedrawRow = RedrawRow + 1
        
        End If
    
        RedrawControls.ControlTabRows(RedrawTab) = RedrawRow
    
    Next

    If RedrawControls.ControlTabsPerRow > 0 Then

        RedrawControls.ControlTotalRows = (RedrawControls.ControlTotalTabs - 1) \ RedrawControls.ControlTabsPerRow + 1

    Else
    
        RedrawVisible = Not RedrawVisible
    
        If RedrawVisible = False Then
        
            RedrawControls.ControlTotalRows = RedrawRow + 1
            
            GoTo loopredrawcalculation
    
        End If
        
    End If
    
    RedrawRow = RedrawControls.ControlTotalRows - 1
    
    ReDim Preserve RedrawControls.ControlRowNumbers(RedrawRow)

    Call SetActiveTab(RedrawMe, RedrawControls, RedrawControls.ControlActiveTab)

End Sub

Public Sub DoActivateTabWithControl(ActivateMe As DataLTab, ActivateControls As DLTControls, ControlHandle As Long)

Dim ActivateTab As Integer
Dim ActivateHandle As String
Dim ActivateFound As Integer

ActivateHandle = "|" + Trim$(Str$(ControlHandle))
ActivateFound = -1

ActivateTab = 0
While ActivateTab < ActivateControls.ControlTotalTabs And ActivateFound = -1

    If InStr("|" + ActivateControls.ControlTabControls(ActivateTab), ActivateHandle) > 0 Then ActivateFound = ActivateTab

    ActivateTab = ActivateTab + 1
Wend

If ActivateFound >= 0 Then Call SetActiveTab(ActivateMe, ActivateControls, ActivateFound)

End Sub

Public Sub ReloadReadData(ReloadMe As DataLTab, ReloadControls As DLTControls)

Dim ReadHandle As String
Dim ReadObject As Control
Dim ReadControlname As String
Dim ReadControlindex As String
Dim ReadControltab As String
Dim ReadPosition As Integer
Dim ReadIndex As Integer
Dim ReadTab As Integer
Dim ReadValue As Long

If ReloadControls.ControlReadData = "" Then Exit Sub

ReloadMe.ContainedControlsCount ReadPosition

If ReadPosition = 0 Then Exit Sub

Debug.Print
Debug.Print "RELOADREADDATA= " & ReloadControls.ControlReadData

While ReloadControls.ControlReadData <> ""

    ReadControltab = ReturnReadPart(ReloadControls.ControlReadData)
    ReadControlname = ReturnReadPart(ReloadControls.ControlReadData)
    ReadControlindex = ReturnReadPart(ReloadControls.ControlReadData)
    ReadTab = Val(ReadControltab)

    If Trim$(Str$(ReadTab)) = ReadControltab And ReadTab >= 0 And ReadTab < ReloadControls.ControlTotalTabs Then

        ReloadMe.ContainedControlsSearch ReadControlname, ReadControlindex, ReadObject

        On Error Resume Next
        ReadValue = -1
        ReadValue = ReadObject.hWnd
        On Error GoTo 0

        If ReadValue > 0 Then

            ReadHandle = "|" + Trim$(Str$(ReadValue)) + "|"

            ReadIndex = 0
            While ReadIndex < ReloadControls.ControlTotalTabs

                ReadPosition = InStr("|" + ReloadControls.ControlTabControls(ReadIndex), ReadHandle)

                If ReadPosition > 0 Then ReloadControls.ControlTabControls(ReadIndex) = Left$(ReloadControls.ControlTabControls(ReadIndex), ReadPosition - 1) + Mid$(ReloadControls.ControlTabControls(ReadIndex), ReadPosition + Len(ReadHandle) - 1)

                ReadIndex = ReadIndex + 1
            Wend

            ReloadControls.ControlTabControls(ReadTab) = ReloadControls.ControlTabControls(ReadTab) + Trim$(Str$(ReadValue)) + "|"
            Debug.Print "FOUNDTAB" & ReadTab & "= " & ReloadControls.ControlTabControls(ReadTab)

        End If
    
    End If

Wend

End Sub

Public Sub LetTotalTabs(LetMe As DataLTab, LetControls As DLTControls, NewTotalTabs)

Dim NewIndex As Integer
Dim NewUpdate As Integer
Dim NewTabhandles As String
Dim NewItemindex As Integer
Dim NewItemtotal As Integer
Dim NewObject As Control
Dim NewHandle As String
Dim NewAbort As Boolean
Dim NewValue As Long

NewUpdate = NewTotalTabs
NewAbort = False

If NewUpdate < 1 Then NewUpdate = 1

If LetControls.ControlActiveTab >= NewUpdate Then

    NewIndex = NewUpdate - 1

    Call SetActiveTab(LetMe, LetControls, NewIndex)

End If

While LetControls.ControlTotalTabs < NewUpdate

    LetControls.ControlTotalTabs = LetControls.ControlTotalTabs + 1

    Call RedimTabArrays(LetControls)

    NewIndex = LetControls.ControlTotalTabs - 1
    
    LetControls.ControlTabCaptions(NewIndex) = ""
    LetControls.ControlTabEnableds(NewIndex) = True
    LetControls.ControlTabVisibles(NewIndex) = True
    LetControls.ControlTabGoodStates(NewIndex) = True
    LetControls.ControlTabControls(NewIndex) = ""

Wend

While LetControls.ControlTotalTabs > NewUpdate And NewAbort = False

    NewIndex = LetControls.ControlTotalTabs - 1

    If LetControls.ControlTabControls(NewIndex) <> "" Then
    
        NewTabhandles = "|" + LetControls.ControlTabControls(NewIndex)
        
        LetMe.ContainedControlsCount NewItemtotal

        NewItemindex = 0
        While NewItemindex < NewItemtotal
    
            SetMe.ContainedControlsItem NewItemindex, SetObject
        
            On Error Resume Next
            NewValue = -1
            NewValue = NewObject.hWnd
            On Error GoTo 0
            
            If NewValue > 0 Then
            
                NewHandle = "|" + Trim$(Str$(NewValue)) + "|"
    
                If InStr(NewTabhandles, NewHandle) > 0 Then NewAbort = True
            
            End If
            
            NewItemindex = NewItemindex + 1
        Wend
    
    End If
    
    If NewAbort = False Then
    
        LetControls.ControlTotalTabs = LetControls.ControlTotalTabs - 1
    
        Call RedimTabArrays(LetControls)
    
    End If

Wend

End Sub

Public Sub LetActiveTab(LetMe As DataLTab, LetControls As DLTControls, NewActiveTab As Integer)

Dim ChangeTab As Integer

ChangeTab = NewActiveTab

If ChangeTab < 0 Then ChangeTab = 0
If ChangeTab >= LetControls.ControlTotalTabs Then ChangeTab = LetControls.ControlTotalTabs - 1

If LetControls.ControlTabVisibles(ChangeTab) = False Then

    LetControls.ControlTabVisibles(ChangeTab) = True
    
    LetControls.ControlActiveTab = ChangeTab
    
    LetMe.RedrawAllTabs

Else

    Call SetActiveTab(LetMe, LetControls, ChangeTab)

End If

End Sub

Public Function LetTabVisible(LetMe As DataLTab, LetControls As DLTControls, TabIndex As Integer, NewTabVisible As Boolean) As Boolean

Dim CheckIndex As Integer
Dim CheckFound As Integer
Dim CheckResult As Boolean

CheckResult = False

If TabIndex >= 0 And TabIndex < LetControls.ControlTotalTabs Then

    CheckFound = -1
    
    If NewTabVisible = False Then
    
        CheckIndex = 0
        While CheckIndex < LetControls.ControlTotalTabs And CheckFound = -1
    
            If LetControls.ControlTabVisibles(CheckIndex) = True Then
                If CheckFound = -1 And CheckIndex <> TabIndex Then CheckFound = CheckIndex
            End If

            CheckIndex = CheckIndex + 1
        Wend
    
    End If
    
    If CheckFound > -1 Or NewTabVisible = True Then
    
        If NewTabVisible = False And TabIndex = LetControls.ControlActiveTab Then Call SetActiveTab(LetMe, LetControls, CheckFound)
        
        LetControls.ControlTabVisibles(TabIndex) = NewTabVisible

        LetMe.RedrawAllTabs

        CheckResult = True

    End If
    
End If

LetTabVisible = CheckResult

End Function

Public Function ReturnTabCaption(ReturnControls As DLTControls, ReturnIndex As Integer) As String

ReturnTabCaption = ReturnControls.ControlTabCaptions(ReturnIndex)

If ReturnTabCaption = "" Then ReturnTabCaption = "Tab" + Trim$(Str$(ReturnIndex))

End Function

Public Function GetObjectIndex(GetControl As Object) As String

If GetControl.Parent.Controls(GetControl.Name) Is GetControl Then
    GetObjectIndex = ""
Else
    GetObjectIndex = GetControl.Index
End If

End Function

Private Function ReturnReadPart(ReadData As String) As String

Dim ReadPosition As Integer
Dim ReadReturn As String

ReadPosition = InStr(ReadData + "|", "|")
ReadReturn = Left$(ReadData, ReadPosition - 1)
ReadData = Mid$(ReadData, ReadPosition + 1)

ReturnReadPart = ReadReturn

End Function

Private Sub RedrawStateBox(RedrawMe As DataLTab, RedrawControls As DLTControls, RedrawLeft As Integer, RedrawTop As Integer, RedrawRight As Integer, RedrawBottom As Integer, RedrawColor As OLE_COLOR)

Dim RedrawOffset As Integer

If (RedrawControls.ControlFlatMode And 2) = 2 Then
    
    RedrawMe.DrawLine RedrawLeft, RedrawTop, RedrawRight, RedrawBottom, RedrawColor, "BF"
    RedrawMe.DrawLine RedrawLeft, RedrawTop, RedrawRight, RedrawBottom, &H646464, "B"

Else
    
    If RedrawControls.ControlOffsetMode = 1 Then
        RedrawOffset = 15
    Else
        RedrawOffset = 0
    End If
    
    RedrawMe.DrawLine RedrawLeft + RedrawOffset, RedrawTop + RedrawOffset, RedrawRight - RedrawOffset, RedrawBottom - RedrawOffset, RedrawColor, "BF"
    RedrawMe.DrawLine RedrawLeft + RedrawOffset, RedrawTop + RedrawOffset, RedrawRight - RedrawOffset, RedrawTop + RedrawOffset, &HA0A0A0
    RedrawMe.DrawLine RedrawLeft + RedrawOffset, RedrawTop + RedrawOffset, RedrawLeft + RedrawOffset, RedrawBottom - RedrawOffset, &HA0A0A0
    RedrawMe.DrawLine RedrawRight - RedrawOffset, RedrawTop + RedrawOffset, RedrawRight - RedrawOffset, RedrawBottom - RedrawOffset, &HFFFFFF
    RedrawMe.DrawLine RedrawLeft + RedrawOffset, RedrawBottom - RedrawOffset, RedrawRight - RedrawOffset + 15, RedrawBottom - RedrawOffset, &HFFFFFF
    RedrawMe.DrawLine RedrawLeft + RedrawOffset + 15, RedrawTop + RedrawOffset + 15, RedrawRight - RedrawOffset - 15, RedrawTop + RedrawOffset + 15, &H696969
    RedrawMe.DrawLine RedrawLeft + RedrawOffset + 15, RedrawTop + RedrawOffset + 15, RedrawLeft + RedrawOffset + 15, RedrawBottom - RedrawOffset - 15, &H696969
    RedrawMe.DrawLine RedrawRight - RedrawOffset - 15, RedrawTop + RedrawOffset + 15, RedrawRight - RedrawOffset - 15, RedrawBottom - RedrawOffset, &HE3E3E3
    RedrawMe.DrawLine RedrawLeft + RedrawOffset + 15, RedrawBottom - RedrawOffset - 15, RedrawRight - RedrawOffset, RedrawBottom - RedrawOffset - 15, &HE3E3E3

End If

End Sub

Private Sub RedrawOneTab(RedrawMe As DataLTab, RedrawControls As DLTControls, RedrawTab As Integer)

Dim RedrawTop As Integer
Dim RedrawLeft As Integer
Dim RedrawRight As Integer
Dim RedrawWidth As Integer
Dim RedrawBottom As Integer
Dim RedrawActive As Integer
Dim RedrawColor As OLE_COLOR
Dim RedrawCurrentY As Integer
Dim RedrawCurrentX As Integer
Dim RedrawCaption As String
Dim RedrawOffset As Integer
Dim RedrawState As Integer
Dim RedrawFrame As Integer
Dim RedrawRow As Integer

If RedrawTab < 0 Or RedrawTab >= RedrawControls.ControlTotalTabs Then Exit Sub
If RedrawControls.ControlTabVisibles(RedrawTab) = False Then Exit Sub

RedrawLeft = RedrawControls.ControlTabLefts(RedrawTab)
RedrawRight = RedrawControls.ControlTabWidths(RedrawTab) + RedrawLeft
RedrawOffset = RedrawControls.ControlOffsetWidth
RedrawState = RedrawControls.ControlStateWidth

If RedrawState < 60 Then RedrawState = 60
If RedrawOffset < 0 Or RedrawControls.ControlOffsetMode = 1 Then RedrawOffset = 0

If (RedrawControls.ControlFlatMode And 2) = 0 And RedrawControls.ControlStateMode > 1 Then
    If RedrawControls.ControlOffsetMode = 1 Then RedrawState = RedrawState + 30
    RedrawState = RedrawState + 30
End If

If RedrawControls.ControlInsideBorder = True Then
    RedrawFrame = 15
Else
    RedrawFrame = 0
End If

If RedrawControls.ControlTotalRows > 1 Then

    RedrawRow = -1
    
    RedrawTop = 0
    While RedrawTop < RedrawControls.ControlTotalRows And RedrawRow = -1
    
        If RedrawControls.ControlRowNumbers(RedrawTop) = RedrawControls.ControlTabRows(RedrawTab) Then RedrawRow = RedrawTop
    
        RedrawTop = RedrawTop + 1
    Wend
    
Else

    RedrawRow = 0

End If

RedrawTop = RedrawRow * RedrawControls.ControlHeightInactive
RedrawBottom = (RedrawControls.ControlTotalRows - 1) * RedrawControls.ControlHeightInactive + RedrawControls.ControlHeightActive

If RedrawTab = RedrawControls.ControlActiveTab Then

    RedrawWidth = RedrawMe.ScaleWidth - 15
    RedrawBottom = RedrawMe.ScaleHeight - 15
    RedrawActive = RedrawTop + RedrawControls.ControlHeightActive
    
    RedrawMe.DrawLine RedrawLeft, RedrawTop, RedrawRight, RedrawActive, RedrawControls.ControlColorActive, "BF"
    RedrawMe.DrawLine 0, RedrawTop + RedrawControls.ControlHeightActive, RedrawWidth, RedrawBottom, RedrawControls.ControlColorActive, "BF"

    If (RedrawControls.ControlFlatMode And 1) = 1 Then
        
        RedrawMe.DrawLine RedrawLeft, RedrawTop, RedrawRight, RedrawTop + RedrawFrame, &H646464, "B"
        RedrawMe.DrawLine RedrawLeft, RedrawTop, RedrawLeft + RedrawFrame, RedrawActive + RedrawFrame, &H646464, "B"
        RedrawMe.DrawLine RedrawRight, RedrawTop, RedrawRight - RedrawFrame, RedrawActive + RedrawFrame, &H646464, "B"
        RedrawMe.DrawLine RedrawLeft + RedrawFrame, RedrawActive, 0, RedrawActive + RedrawFrame, &H646464, "B"
        RedrawMe.DrawLine RedrawRight - RedrawFrame, RedrawActive, RedrawWidth, RedrawActive + RedrawFrame, &H646464, "B"
        
        RedrawMe.DrawLine 0, RedrawActive, RedrawFrame, RedrawBottom, &H646464, "B"
        RedrawMe.DrawLine RedrawWidth, RedrawActive, RedrawWidth - RedrawFrame, RedrawBottom, &H646464, "B"
        RedrawMe.DrawLine 0, RedrawBottom, RedrawWidth, RedrawBottom - RedrawFrame, &H646464, "B"
    
    Else
        
        RedrawMe.DrawLine RedrawLeft, RedrawTop, RedrawRight, RedrawTop, &H646464, "B"
        RedrawMe.DrawLine RedrawLeft + 15, RedrawTop + 15, RedrawRight - 30, RedrawTop + 15, &HFFFFFF, "B"
        RedrawMe.DrawLine RedrawLeft + 15, RedrawTop + 30, RedrawRight - 45, RedrawTop + 30, &HFFFFFF, "B"
        RedrawMe.DrawLine RedrawLeft, RedrawTop, RedrawLeft, RedrawActive, &H646464, "B"
        RedrawMe.DrawLine RedrawLeft + 15, RedrawTop + 15, RedrawLeft + 15, RedrawActive + 15, &HFFFFFF, "B"
        RedrawMe.DrawLine RedrawLeft + 30, RedrawTop + 15, RedrawLeft + 30, RedrawActive + 30, &HFFFFFF, "B"
        RedrawMe.DrawLine RedrawRight, RedrawTop, RedrawRight, RedrawActive, &H646464, "B"
        RedrawMe.DrawLine RedrawRight - 15, RedrawTop + 15, RedrawRight - 15, RedrawActive, &HA0A0A0, "B"
        RedrawMe.DrawLine RedrawRight - 30, RedrawTop + 30, RedrawRight - 30, RedrawActive + 15, &HA0A0A0, "B"
        RedrawMe.DrawLine 0, RedrawActive, RedrawLeft, RedrawActive, &H646464, "B"
        RedrawMe.DrawLine 15, RedrawActive + 15, RedrawLeft + 15, RedrawActive + 15, &HFFFFFF, "B"
        RedrawMe.DrawLine 15, RedrawActive + 30, RedrawLeft + 30, RedrawActive + 30, &HFFFFFF, "B"
        
        If RedrawRight <> RedrawWidth Then
            RedrawMe.DrawLine RedrawRight, RedrawActive, RedrawWidth, RedrawActive, &H646464, "B"
            RedrawMe.DrawLine RedrawRight - 15, RedrawActive + 15, RedrawWidth - 30, RedrawActive + 15, &HFFFFFF, "B"
            RedrawMe.DrawLine RedrawRight - 30, RedrawActive + 30, RedrawWidth - 45, RedrawActive + 30, &HFFFFFF, "B"
        End If
        
        RedrawMe.DrawLine 30, RedrawActive + 15, 30, RedrawBottom - 30, &HFFFFFF, "B"
        RedrawMe.DrawLine 15, RedrawActive + 15, 15, RedrawBottom - 15, &HFFFFFF, "B"
        RedrawMe.DrawLine 0, RedrawActive, 0, RedrawBottom, &H646464, "B"
        RedrawMe.DrawLine RedrawWidth - 30, RedrawActive + 30, RedrawWidth - 30, RedrawBottom - 15, &HA0A0A0, "B"
        RedrawMe.DrawLine RedrawWidth - 15, RedrawActive + 15, RedrawWidth - 15, RedrawBottom - 15, &HA0A0A0, "B"
        RedrawMe.DrawLine RedrawWidth, RedrawActive, RedrawWidth, RedrawBottom, &H646464, "B"
        RedrawMe.DrawLine 45, RedrawBottom - 30, RedrawWidth - 15, RedrawBottom - 30, &HA0A0A0, "B"
        RedrawMe.DrawLine 30, RedrawBottom - 15, RedrawWidth - 15, RedrawBottom - 15, &HA0A0A0, "B"
        RedrawMe.DrawLine 0, RedrawBottom, RedrawWidth, RedrawBottom, &H646464, "B"
    
        If RedrawControls.ControlInsideBorder = True Then
            
            RedrawMe.DrawLine RedrawLeft + 15, RedrawTop + 45, RedrawRight - 60, RedrawTop + 45, &HFFFFFF, "B"
            RedrawMe.DrawLine RedrawLeft + 45, RedrawTop + 15, RedrawLeft + 45, RedrawActive + 45, &HFFFFFF, "B"
            RedrawMe.DrawLine RedrawRight - 45, RedrawTop + 45, RedrawRight - 45, RedrawActive + 30, &HA0A0A0, "B"
            RedrawMe.DrawLine 15, RedrawActive + 45, RedrawLeft + 45, RedrawActive + 45, &HFFFFFF, "B"
            
            If RedrawRight <> RedrawWidth Then RedrawMe.DrawLine RedrawRight - 45, RedrawActive + 45, RedrawWidth - 60, RedrawActive + 45, &HFFFFFF, "B"
            
            RedrawMe.DrawLine 45, RedrawActive + 15, 45, RedrawBottom - 45, &HFFFFFF, "B"
            RedrawMe.DrawLine RedrawWidth - 45, RedrawActive + 45, RedrawWidth - 45, RedrawBottom - 15, &HA0A0A0, "B"
            RedrawMe.DrawLine 60, RedrawBottom - 45, RedrawWidth - 15, RedrawBottom - 45, &HA0A0A0, "B"
        
            RedrawFrame = 30
        
        Else
    
            RedrawFrame = 15

        End If
    
    End If
    
Else

    If RedrawRow = RedrawControls.ControlTotalRows - 1 Then RedrawTop = RedrawTop + RedrawControls.ControlHeightActive - RedrawControls.ControlHeightInactive
    
    If RedrawControls.ControlTabEnableds(RedrawTab) = False Then
        RedrawColor = RedrawControls.ControlColorDisabled
    Else
        RedrawColor = RedrawControls.ControlColorInactive
    End If
    
    RedrawMe.DrawLine RedrawLeft, RedrawTop, RedrawRight, RedrawBottom - 15, RedrawColor, "BF"
    
    If (RedrawControls.ControlFlatMode And 1) = 1 Then
        
        RedrawMe.DrawLine RedrawLeft, RedrawTop, RedrawRight, RedrawBottom, &H646464, "B"
    
        RedrawFrame = 0
    
    Else
    
        RedrawMe.DrawLine RedrawLeft, RedrawTop, RedrawRight, RedrawTop, &H646464
        RedrawMe.DrawLine RedrawLeft + 15, RedrawTop + 15, RedrawRight - 15, RedrawTop + 15, &HFFFFFF
        RedrawMe.DrawLine RedrawLeft, RedrawTop, RedrawLeft, RedrawBottom, &H646464
        RedrawMe.DrawLine RedrawLeft + 15, RedrawTop + 15, RedrawLeft + 15, RedrawBottom, &HFFFFFF
        RedrawMe.DrawLine RedrawRight, RedrawTop, RedrawRight, RedrawBottom, &H646464
        RedrawMe.DrawLine RedrawRight - 15, RedrawTop + 15, RedrawRight - 15, RedrawBottom, &HA0A0A0
    
        RedrawFrame = 15
    
    End If
    
End If

RedrawTop = RedrawTop + RedrawFrame
RedrawCurrentY = RedrawTop + RedrawControls.ControlCaptionOffset
RedrawTop = RedrawTop + RedrawOffset

If RedrawControls.ControlTabGoodStates(RedrawTab) = True Then
    RedrawColor = RedrawControls.ControlColorValid
Else
    RedrawColor = RedrawControls.ControlColorInvalid
End If

RedrawBottom = RedrawTop + RedrawControls.ControlHeightInactive - RedrawOffset * 2 - 15

If RedrawTab = RedrawControls.ControlActiveTab And (RedrawControls.ControlFlatMode And 1) = 0 Then
    RedrawCurrentY = RedrawCurrentY + 15
    RedrawOffset = RedrawOffset + 15
    RedrawBottom = RedrawBottom + 15
    RedrawTop = RedrawTop + 15
End If

Select Case RedrawControls.ControlStateMode
    
    Case 1
        
        RedrawLeft = RedrawLeft + RedrawOffset + RedrawFrame
        RedrawRight = RedrawRight - RedrawOffset - RedrawFrame
        
        Call RedrawStateBox(RedrawMe, RedrawControls, RedrawLeft, RedrawTop, RedrawRight, RedrawBottom, RedrawColor)
        
        RedrawCurrentX = RedrawLeft + DLTCaptionSpacing
    
    Case 2
        
        RedrawCurrentX = RedrawLeft + DLTCaptionSpacing + RedrawState + RedrawOffset + RedrawFrame
    
    Case Else
        
        RedrawCurrentX = RedrawLeft + DLTCaptionSpacing + RedrawFrame

End Select

RedrawCaption = ReturnTabCaption(RedrawControls, RedrawTab)

If RedrawControls.ControlTabsPerRow > 0 Then

    RedrawWidth = RedrawControls.ControlTabWidths(RedrawTab) - DLTCaptionSpacing * 2
    
    If RedrawControls.ControlInsideBorder = True Then RedrawWidth = RedrawWidth - 30
    
    Select Case RedrawControls.ControlStateMode
        Case 1:    RedrawWidth = RedrawWidth - RedrawOffset * 2
        Case 2, 3: RedrawWidth = RedrawWidth - RedrawState - RedrawOffset
    End Select
    
    While RedrawCaption <> "" And RedrawMe.TextWidth(RedrawCaption) > RedrawWidth
        RedrawCaption = Left$(RedrawCaption, Len(RedrawCaption) - 1)
    Wend

End If

RedrawMe.DrawText RedrawControls.ControlColorCaption, RedrawCurrentX, RedrawCurrentY, RedrawCaption
      
Select Case RedrawControls.ControlStateMode
    
    Case 2
        
        RedrawLeft = RedrawLeft + RedrawOffset + RedrawFrame
    
    Case 3
        
        RedrawLeft = RedrawRight - RedrawState - RedrawOffset - RedrawFrame
    
    Case Else
        
        Exit Sub

End Select

RedrawRight = RedrawLeft + RedrawState

Call RedrawStateBox(RedrawMe, RedrawControls, RedrawLeft, RedrawTop, RedrawRight, RedrawBottom, RedrawColor)

End Sub

Private Sub SetActiveTab(SetMe As DataLTab, SetControls As DLTControls, SetNewtab As Integer)

Dim SetOldtab As Integer
Dim SetObject As Control
Dim SetIndexrow As Integer
Dim SetIndextab As Integer
Dim SetFoundtab As Integer
Dim SetTabcontrols() As String
Dim SetItemindex As Integer
Dim SetItemtotal As Integer
Dim SetHandle As String
Dim SetValue As Long

If SetControls.ControlActiveTab < 0 Then Exit Sub
If SetControls.ControlTotalTabs < 1 Then Exit Sub

Call ReloadReadData(SetMe, SetControls)

Debug.Print

SetOldtab = SetControls.ControlActiveTab
SetIndextab = SetControls.ControlTotalTabs - 1

ReDim Preserve SetTabcontrols(SetIndextab)

SetIndextab = 0
While SetIndextab < SetControls.ControlTotalTabs
    SetTabcontrols(SetIndextab) = "|"
    SetIndextab = SetIndextab + 1
Wend

SetMe.ContainedControlsCount SetItemtotal

SetItemindex = 0
While SetItemindex < SetItemtotal
    
    SetMe.ContainedControlsItem SetItemindex, SetObject

    On Error Resume Next
    SetValue = -1
    SetValue = SetObject.hWnd
    On Error GoTo 0
    
    If SetValue > 0 Then
    
        SetHandle = "|" + Trim$(Str$(SetValue)) + "|"
    
        If SetObject.Left >= 0 Then
    
            SetFoundtab = SetOldtab
    
        Else
    
            SetIndextab = 0
            SetFoundtab = -1
            While SetIndextab < SetControls.ControlTotalTabs And SetFoundtab = -1
                If InStr("|" + SetControls.ControlTabControls(SetIndextab), SetHandle) > 0 Then SetFoundtab = SetIndextab
                SetIndextab = SetIndextab + 1
            Wend
    
            If SetFoundtab = -1 Then SetFoundtab = SetOldtab

        End If
    
        SetTabcontrols(SetFoundtab) = SetTabcontrols(SetFoundtab) + Mid$(SetHandle, 2)
        
        If SetFoundtab = SetNewtab Then
    
            If SetObject.Left < 0 Then
                SetObject.Left = SetObject.Left + 75000
                Debug.Print "SHOW= " & SetObject.hWnd & " " & SetObject.Name
            End If
    
        Else
        
            If SetObject.Left >= 0 Then
                SetObject.Left = SetObject.Left - 75000
                Debug.Print "HIDE= " & SetObject.hWnd & " " & SetObject.Name
            End If
    
        End If

    End If
    
    SetItemindex = SetItemindex + 1
Wend
    
SetIndextab = 0
While SetIndextab < SetControls.ControlTotalTabs
    
    SetControls.ControlTabControls(SetIndextab) = Mid$(SetTabcontrols(SetIndextab), 2)
    Debug.Print "TAB" & SetIndextab & "-AFTER=" & SetControls.ControlTabControls(SetIndextab)
    
    SetIndextab = SetIndextab + 1
Wend

SetControls.ControlActiveTab = SetNewtab

SetIndextab = -1
SetFoundtab = SetControls.ControlTabRows(SetNewtab)
SetIndexrow = SetControls.ControlTotalRows - 1
    
SetControls.ControlRowNumbers(SetIndexrow) = SetFoundtab
    
While SetIndexrow > 0
    
    SetIndexrow = SetIndexrow - 1
    SetIndextab = SetIndextab + 1
        
    If SetIndextab = SetFoundtab Then SetIndextab = SetIndextab + 1
        
    SetControls.ControlRowNumbers(SetIndexrow) = SetIndextab
    
Wend

SetMe.DrawCls SetControls.ControlColorBack

SetIndexrow = 0
While SetIndexrow < SetControls.ControlTotalRows
    
    SetIndextab = 0
    While SetIndextab < SetControls.ControlTotalTabs
        
        If SetControls.ControlTabRows(SetIndextab) = SetControls.ControlRowNumbers(SetIndexrow) Then Call RedrawOneTab(SetMe, SetControls, SetIndextab)

        SetIndextab = SetIndextab + 1
    Wend

    SetIndexrow = SetIndexrow + 1
Wend
    
If SetOldtab = SetNewtab Then Exit Sub

SetMe.RaiseAfterTabChange SetOldtab

End Sub

Private Sub RedimTabArrays(RedimControls As DLTControls)

Dim RedimSize As Integer

RedimSize = RedimControls.ControlTotalTabs - 1

ReDim Preserve RedimControls.ControlTabCaptions(RedimSize)
ReDim Preserve RedimControls.ControlTabEnableds(RedimSize)
ReDim Preserve RedimControls.ControlTabVisibles(RedimSize)
ReDim Preserve RedimControls.ControlTabGoodStates(RedimSize)
ReDim Preserve RedimControls.ControlTabWidths(RedimSize)
ReDim Preserve RedimControls.ControlTabLefts(RedimSize)
ReDim Preserve RedimControls.ControlTabRows(RedimSize)
ReDim Preserve RedimControls.ControlTabControls(RedimSize)

End Sub


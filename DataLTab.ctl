VERSION 5.00
Begin VB.UserControl DataLTab 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   PropertyPages   =   "DataLTab.ctx":0000
   ScaleHeight     =   1995
   ScaleWidth      =   5550
   ToolboxBitmap   =   "DataLTab.ctx":0019
End
Attribute VB_Name = "DataLTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"

Public Enum DLTStateMode
    WithoutState = 0
    WithStateFullSize = 1
    WithStateLeft = 2
    WithStateRight = 3
End Enum
Public Enum DLTOffsetMode
    StateInTabOffset = 0
    TabInControlOffset = 1
    BothOffset = 2
End Enum
Public Enum DLTFlatMode
    NotFlat = 0
    FlatBorders = 1
    FlatState = 2
    FlatAll = 3
End Enum

Dim DLTControls As DLTControls

Dim WithEvents DLTCaptionFont As StdFont
Attribute DLTCaptionFont.VB_VarHelpID = -1

Event AfterTabChange(PreviousTab As Integer)

Private Sub UserControl_Initialize()

Call InitializeDLTControls(DLTControls)

End Sub

Private Sub UserControl_Show()

Call RedrawAllTabs

End Sub

Private Sub UserControl_Paint()

Call ReloadReadData(Me, DLTControls)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

Set DLTCaptionFont = PropBag.ReadProperty("CaptionFont", Ambient.Font)

Call ReadDLTPropertyBag(DLTControls, PropBag)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

Call ReloadReadData(Me, DLTControls)

Call PropBag.WriteProperty("CaptionFont", DLTCaptionFont, Ambient.Font)

Call WriteDLTPropertyBag(Me, DLTControls, PropBag)

End Sub

Private Sub UserControl_InitProperties()

Call InitPropertiesDLTControls(DLTControls)

Set DLTCaptionFont = Ambient.Font

End Sub

Private Sub UserControl_Resize()

Call ReloadReadData(Me, DLTControls)

Call RedrawAllTabs

End Sub

Private Sub UserControl_MouseDown(MouseButton As Integer, MouseShift As Integer, MouseX As Single, MouseY As Single)

Call DoMouseMove(Me, DLTControls, MouseButton, MouseShift, MouseX, MouseY)

End Sub

Public Property Get hWnd() As Long

hWnd = UserControl.hWnd

End Property

Public Property Get TabVisible(TabIndex As Integer) As Boolean

If TabIndex < 0 Or TabIndex >= DLTControls.ControlTotalTabs Then
    TabVisible = False
Else
    TabVisible = DLTControls.ControlTabVisibles(TabIndex)
End If

End Property

Public Property Let TabVisible(TabIndex As Integer, NewTabVisible As Boolean)

If LetTabVisible(Me, DLTControls, TabIndex, NewTabVisible) = True Then PropertyChanged "TabVisible"

End Property

Public Property Get FlatMode() As DLTFlatMode

FlatMode = DLTControls.ControlFlatMode

End Property

Public Property Let FlatMode(ByVal NewFlatMode As DLTFlatMode)

If NewFlatMode <> DLTControls.ControlFlatMode Then

    DLTControls.ControlFlatMode = NewFlatMode

    Call RedrawAllTabs

End If

End Property

Public Property Get InsideBorder() As Boolean

InsideBorder = DLTControls.ControlInsideBorder

End Property

Public Property Let InsideBorder(ByVal NewInsideBorder As Boolean)

If NewInsideBorder <> DLTControls.ControlInsideBorder Then

    DLTControls.ControlInsideBorder = NewInsideBorder

    Call RedrawAllTabs

End If

End Property

Public Property Get ActiveTab() As Integer

ActiveTab = DLTControls.ControlActiveTab

End Property

Public Property Let ActiveTab(ByVal NewActiveTab As Integer)

Call LetActiveTab(Me, DLTControls, NewActiveTab)

PropertyChanged "ActiveTab"

End Property

Public Property Get TabCaption(TabIndex As Integer) As String

If TabIndex < 0 Or TabIndex >= DLTControls.ControlTotalTabs Then
    TabCaption = ""
Else
    TabCaption = ReturnTabCaption(DLTControls, TabIndex)
End If

End Property

Public Property Let TabCaption(TabIndex As Integer, NewTabCaption As String)

If TabIndex >= 0 And TabIndex < DLTControls.ControlTotalTabs Then

    DLTControls.ControlTabCaptions(TabIndex) = NewTabCaption

    Call RedrawAllTabs

    PropertyChanged "TabCaption"

End If

End Property

Public Property Get TabEnabled(TabIndex As Integer) As Boolean

If TabIndex < 0 Or TabIndex >= DLTControls.ControlTotalTabs Then
    TabEnabled = False
Else
    TabEnabled = DLTControls.ControlTabEnableds(TabIndex)
End If

End Property

Public Property Let TabEnabled(TabIndex As Integer, NewTabEnabled As Boolean)

If TabIndex >= 0 And TabIndex < DLTControls.ControlTotalTabs Then

    DLTControls.ControlTabEnableds(TabIndex) = NewTabEnabled

    Call RedrawAllTabs

    PropertyChanged "TabEnabled"

End If

End Property

Public Property Get TabGoodState(TabIndex As Integer) As Boolean

If TabIndex < 0 Or TabIndex >= DLTControls.ControlTotalTabs Then
    TabGoodState = False
Else
    TabGoodState = DLTControls.ControlTabGoodStates(TabIndex)
End If

End Property

Public Property Let TabGoodState(TabIndex As Integer, NewTabGoodState As Boolean)

If TabIndex >= 0 And TabIndex < DLTControls.ControlTotalTabs Then

    DLTControls.ControlTabGoodStates(TabIndex) = NewTabGoodState

    Call RedrawAllTabs

    PropertyChanged "TabGoodState"

End If

End Property

Public Property Get TotalTabs() As Integer
Attribute TotalTabs.VB_ProcData.VB_Invoke_Property = "PropertyPage1"

TotalTabs = DLTControls.ControlTotalTabs

End Property

Public Property Let TotalTabs(ByVal NewTotalTabs As Integer)

If Ambient.UserMode = True Then Exit Property

Call LetTotalTabs(Me, DLTControls, NewTotalTabs)

Call RedrawAllTabs

PropertyChanged "TotalTabs"

End Property

Public Property Let TabsPerRow(ByVal NewTabsPerRow As Long)

DLTControls.ControlTabsPerRow = NewTabsPerRow

If DLTControls.ControlTabsPerRow < 0 Then DLTControls.ControlTabsPerRow = 0

Call RedrawAllTabs

PropertyChanged "TabsPerRow"

End Property

Public Property Get TabsPerRow() As Long
    
TabsPerRow = DLTControls.ControlTabsPerRow

End Property

Public Property Let OffsetWidth(ByVal NewOffsetWidth As Long)

DLTControls.ControlOffsetWidth = NewOffsetWidth

If DLTControls.ControlOffsetWidth < 0 Then DLTControls.ControlOffsetWidth = 0

Call RedrawAllTabs

End Property

Public Property Get OffsetWidth() As Long
    
OffsetWidth = DLTControls.ControlOffsetWidth

End Property

Public Property Let StateWidth(ByVal NewStateWidth As Long)

DLTControls.ControlStateWidth = NewStateWidth

If DLTControls.ControlStateWidth < 60 Then DLTControls.ControlStateWidth = 60

Call RedrawAllTabs

End Property

Public Property Get StateWidth() As Long
    
StateWidth = DLTControls.ControlStateWidth

End Property

Public Property Let DisabledColor(ByVal NewDisabledColor As OLE_COLOR)

DLTControls.ControlColorDisabled = NewDisabledColor

Call RedrawAllTabs

End Property

Public Property Get DisabledColor() As OLE_COLOR

DisabledColor = DLTControls.ControlColorDisabled

End Property

Public Property Let InactiveColor(ByVal NewInactiveColor As OLE_COLOR)

DLTControls.ControlColorInactive = NewInactiveColor

Call RedrawAllTabs

End Property

Public Property Get InactiveColor() As OLE_COLOR

InactiveColor = DLTControls.ControlColorInactive

End Property

Public Property Let ValidColor(ByVal NewValidColor As OLE_COLOR)

DLTControls.ControlColorValid = NewValidColor

Call RedrawAllTabs

End Property

Public Property Get ValidColor() As OLE_COLOR

ValidColor = DLTControls.ControlColorValid

End Property

Public Property Let InvalidColor(ByVal NewInvalidColor As OLE_COLOR)

DLTControls.ControlColorInvalid = NewInvalidColor

Call RedrawAllTabs

End Property

Public Property Get InvalidColor() As OLE_COLOR

InvalidColor = DLTControls.ControlColorInvalid

End Property

Public Property Let ActiveColor(ByVal NewActiveColor As OLE_COLOR)

DLTControls.ControlColorActive = NewActiveColor

Call RedrawAllTabs

End Property

Public Property Get ActiveColor() As OLE_COLOR

ActiveColor = DLTControls.ControlColorActive

End Property

Public Property Let BackgroundColor(ByVal NewBackgroundColor As OLE_COLOR)

DLTControls.ControlColorBack = NewBackgroundColor

Call RedrawAllTabs

End Property

Public Property Get BackgroundColor() As OLE_COLOR

BackgroundColor = DLTControls.ControlColorBack

End Property

Public Property Let CaptionColor(ByVal NewCaptionColor As OLE_COLOR)

DLTControls.ControlColorCaption = NewCaptionColor

Call RedrawAllTabs

End Property

Public Property Get CaptionColor() As OLE_COLOR

CaptionColor = DLTControls.ControlColorCaption

End Property

Public Property Let ActiveHeight(ByVal NewActiveHeight As Integer)

DLTControls.ControlHeightActive = NewActiveHeight

Call RedrawAllTabs

End Property

Public Property Get ActiveHeight() As Integer

ActiveHeight = DLTControls.ControlHeightActive

End Property

Public Property Let InactiveHeight(ByVal NewInactiveHeight As Integer)

DLTControls.ControlHeightInactive = NewInactiveHeight

Call RedrawAllTabs

End Property

Public Property Get InactiveHeight() As Integer

InactiveHeight = DLTControls.ControlHeightInactive

End Property

Public Property Let OffsetMode(ByVal NewOffsetMode As DLTOffsetMode)

DLTControls.ControlOffsetMode = NewOffsetMode

Call RedrawAllTabs

End Property

Public Property Get OffsetMode() As DLTOffsetMode

OffsetMode = DLTControls.ControlOffsetMode

End Property

Public Property Let StateMode(ByVal NewStateMode As DLTStateMode)

DLTControls.ControlStateMode = NewStateMode

Call RedrawAllTabs

End Property

Public Property Get StateMode() As DLTStateMode

StateMode = DLTControls.ControlStateMode

End Property

Public Property Get CaptionFont() As StdFont

Set CaptionFont = DLTCaptionFont

End Property

Public Property Set CaptionFont(ByVal NewCaptionFont As StdFont)

Set DLTCaptionFont = NewCaptionFont

Call RedrawAllTabs

End Property

Public Sub ActivateTabWithControl(ControlHandle As Long)

Call DoActivateTabWithControl(Me, DLTControls, ControlHandle)

End Sub

Friend Function TextWidth(TextString As String) As Long

TextWidth = UserControl.TextWidth(TextString)

End Function

Friend Function TextHeight(TextString As String) As Long

TextHeight = UserControl.TextHeight(TextString)

End Function

Friend Function ScaleWidth() As Long

ScaleWidth = UserControl.ScaleWidth

End Function

Friend Function ScaleHeight() As Long

ScaleHeight = UserControl.ScaleHeight

End Function

Friend Sub ContainedControlsCount(TotalCount As Integer)

On Error Resume Next

TotalCount = 0
TotalCount = UserControl.ContainedControls.Count

On Error GoTo 0

End Sub

Friend Sub ContainedControlsSearch(SearchName As String, SearchIndex As String, SearchFound As Control)

Dim SearchObject As Control

Set SearchFound = Nothing

For Each SearchObject In UserControl.ContainedControls
    If SearchObject.Name = SearchName Then
        If GetObjectIndex(SearchObject) = SearchIndex Then

            Set SearchFound = SearchObject
            
            Exit Sub
            
        End If
    End If
Next

End Sub

Friend Sub ContainedControlsItem(ItemNumber As Integer, ItemFound As Control)

Set ItemFound = Nothing

On Error Resume Next

Set ItemFound = UserControl.ContainedControls.Item(ItemNumber)

On Error GoTo 0

End Sub

Friend Sub DrawCls(BackgroundColor As OLE_COLOR)

UserControl.Cls
UserControl.ScaleMode = 1
UserControl.BackColor = BackgroundColor

End Sub

Friend Sub DrawText(DrawColor As OLE_COLOR, DrawX As Integer, DrawY As Integer, DrawText As String)

UserControl.ForeColor = DrawColor
UserControl.CurrentX = DrawX
UserControl.CurrentY = DrawY
UserControl.Print DrawText

End Sub

Friend Sub DrawLine(DrawX1 As Integer, DrawY1 As Integer, DrawX2 As Integer, DrawY2 As Integer, DrawColor As OLE_COLOR, Optional DrawMode As String = "")

Select Case DrawMode
    Case "B":  UserControl.Line (DrawX1, DrawY1)-(DrawX2, DrawY2), DrawColor, B
    Case "BF": UserControl.Line (DrawX1, DrawY1)-(DrawX2, DrawY2), DrawColor, BF
    Case Else: UserControl.Line (DrawX1, DrawY1)-(DrawX2, DrawY2), DrawColor
End Select

End Sub

Friend Sub RedrawAllTabs()

Set UserControl.Font = DLTCaptionFont

Call DoRedrawAllTabs(Me, DLTControls)

End Sub

Friend Sub RaiseAfterTabChange(PreviousTab As Integer)

RaiseEvent AfterTabChange(PreviousTab)

End Sub

Private Sub DLTCaptionFont_FontChanged(ByVal PropertyName As String)

Call RedrawAllTabs

End Sub

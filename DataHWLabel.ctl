VERSION 5.00
Begin VB.UserControl DataHWLabel 
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2835
   ScaleHeight     =   270
   ScaleWidth      =   2835
   ToolboxBitmap   =   "DataHWLabel.ctx":0000
   Begin VB.Label HWLabelControl 
      Caption         =   "Label1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "DataHWLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Enum EnumAlignment
    LeftJustify = 0
    RightJustify = 1
    Center = 2
End Enum
Public Enum EnumAppearance
    Flat = 0
    As3D = 1
End Enum
Public Enum EnumBackStyle
    Transparent = 0
    Opaque = 1
End Enum
Public Enum EnumBorderStyle
    None = 0
    FixedSingle = 1
End Enum

Dim WithEvents ControlFont As StdFont
Attribute ControlFont.VB_VarHelpID = -1

Private Sub UserControl_Initialize()

Set ControlFont = HWLabelControl.Font

HWLabelControl.Caption = "HWLabelControl"

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

HWLabelControl.Enabled = PropBag.ReadProperty("Enabled", True)
HWLabelControl.WordWrap = PropBag.ReadProperty("WordWrap", False)
HWLabelControl.Visible = PropBag.ReadProperty("Visible", True)
HWLabelControl.UseMnemonic = PropBag.ReadProperty("UseMnemonic", True)

HWLabelControl.Caption = PropBag.ReadProperty("Caption", "")
HWLabelControl.Tag = PropBag.ReadProperty("Tag", "")

HWLabelControl.Alignment = PropBag.ReadProperty("Alignment", 0)
HWLabelControl.Appearance = PropBag.ReadProperty("Appearance", 1)
HWLabelControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
HWLabelControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)

HWLabelControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
HWLabelControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)

Set ControlFont = PropBag.ReadProperty("Font", Ambient.Font)
Set HWLabelControl.Font = ControlFont

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

Call PropBag.WriteProperty("Enabled", HWLabelControl.Enabled, True)
Call PropBag.WriteProperty("WordWrap", HWLabelControl.WordWrap, False)
Call PropBag.WriteProperty("Visible", HWLabelControl.Visible, True)
Call PropBag.WriteProperty("UseMnemonic", HWLabelControl.UseMnemonic, True)

Call PropBag.WriteProperty("Caption", HWLabelControl.Caption, "")
Call PropBag.WriteProperty("Tag", HWLabelControl.Tag, "")

Call PropBag.WriteProperty("Alignment", HWLabelControl.Alignment, 0)
Call PropBag.WriteProperty("Appearance", HWLabelControl.Appearance, 1)
Call PropBag.WriteProperty("BackStyle", HWLabelControl.BackStyle, 1)
Call PropBag.WriteProperty("BorderStyle", HWLabelControl.BorderStyle, 0)

Call PropBag.WriteProperty("ForeColor", HWLabelControl.ForeColor, &H80000012)
Call PropBag.WriteProperty("BackColor", HWLabelControl.BackColor, &H8000000F)

Call PropBag.WriteProperty("Font", ControlFont, Ambient.Font)

End Sub

Private Sub UserControl_Resize()

HWLabelControl.Width = UserControl.Width
HWLabelControl.Height = HWLabelControl.Height

End Sub

Property Get hWnd() As Long

hWnd = UserControl.hWnd

End Property

Property Get Visible() As Boolean

Visible = HWLabelControl.Visible

End Property

Property Let Visible(NewVisible As Boolean)

HWLabelControl.Visible = NewVisible

End Property

Property Get Enabled() As Boolean

Enabled = HWLabelControl.Enabled

End Property

Property Let Enabled(NewEnabled As Boolean)

HWLabelControl.Enabled = NewEnabled

End Property

Property Get WordWrap() As Boolean

WordWrap = HWLabelControl.WordWrap

End Property

Property Let WordWrap(NewWordWrap As Boolean)

HWLabelControl.WordWrap = NewWordWrap

End Property

Property Get UseMnemonic() As Boolean

UseMnemonic = HWLabelControl.UseMnemonic

End Property

Property Let UseMnemonic(NewUseMnemonic As Boolean)

HWLabelControl.UseMnemonic = NewUseMnemonic

End Property

Property Get Tag() As String

Tag = HWLabelControl.Tag

End Property

Property Let Tag(NewTag As String)

HWLabelControl.Tag = NewTag

End Property

Property Get Caption() As String

Caption = HWLabelControl.Caption

End Property

Property Let Caption(NewCaption As String)

HWLabelControl.Caption = NewCaption

End Property

Property Get BackColor() As OLE_COLOR

BackColor = HWLabelControl.BackColor

End Property

Property Let BackColor(NewBackColor As OLE_COLOR)

HWLabelControl.BackColor = NewBackColor

End Property

Property Get ForeColor() As OLE_COLOR

ForeColor = HWLabelControl.ForeColor

End Property

Property Let ForeColor(NewForeColor As OLE_COLOR)

HWLabelControl.ForeColor = NewForeColor

End Property

Property Get Alignment() As EnumAlignment

Alignment = HWLabelControl.Alignment

End Property

Property Let Alignment(NewAlignment As EnumAlignment)

HWLabelControl.Alignment = NewAlignment

End Property

Property Get Appearance() As EnumAppearance

Appearance = HWLabelControl.Appearance

End Property

Property Let Appearance(NewAppearance As EnumAppearance)

HWLabelControl.Appearance = NewAppearance

End Property

Property Get BackStyle() As EnumBackStyle

BackStyle = HWLabelControl.BackStyle

End Property

Property Let BackStyle(NewBackStyle As EnumBackStyle)

HWLabelControl.BackStyle = NewBackStyle

End Property

Property Get BorderStyle() As EnumBorderStyle

BorderStyle = HWLabelControl.BorderStyle

End Property

Property Let BorderStyle(NewBorderStyle As EnumBorderStyle)

HWLabelControl.BorderStyle = NewBorderStyle

End Property

Public Property Get Font() As StdFont

Set Font = ControlFont

End Property

Public Property Set Font(ByVal NewFont As StdFont)

Set ControlFont = NewFont
Set HWLabelControl.Font = NewFont

End Property

Private Sub ControlFont_FontChanged(ByVal PropertyName As String)

Set HWLabelControl.Font = ControlFont

End Sub

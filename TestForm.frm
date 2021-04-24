VERSION 5.00
Object = "*\AControlProject.vbp"
Begin VB.Form TestForm 
   Caption         =   "LightTab"
   ClientHeight    =   2895
   ClientLeft      =   3885
   ClientTop       =   3645
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   8655
   Begin VB.CommandButton Command11 
      Caption         =   "FlatMode"
      Height          =   495
      Left            =   5040
      TabIndex        =   23
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "InsideBorder"
      Height          =   495
      Left            =   6120
      TabIndex        =   22
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "TabsPerRow"
      Height          =   495
      Left            =   6120
      TabIndex        =   21
      Top             =   1740
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "OffsetMode"
      Height          =   495
      Left            =   6120
      TabIndex        =   20
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Visible 6"
      Height          =   375
      Index           =   6
      Left            =   7560
      TabIndex        =   19
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Visible 5"
      Height          =   375
      Index           =   5
      Left            =   7560
      TabIndex        =   18
      Top             =   2020
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Visible 4"
      Height          =   375
      Index           =   4
      Left            =   7560
      TabIndex        =   10
      Top             =   1640
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Visible 3"
      Height          =   375
      Index           =   3
      Left            =   7560
      TabIndex        =   9
      Top             =   1260
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Visible 2"
      Height          =   375
      Index           =   2
      Left            =   7560
      TabIndex        =   8
      Top             =   880
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Visible 1"
      Height          =   375
      Index           =   1
      Left            =   7560
      TabIndex        =   7
      Top             =   500
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Visible 0"
      Height          =   375
      Index           =   0
      Left            =   7560
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "StateMode"
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin ControlProject.DataLTab LightTab1 
      Height          =   2655
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4683
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      InactiveColor   =   12632256
      DisabledColor   =   8421504
      ActiveHeight    =   330
      InactiveHeight  =   285
      TotalTabs       =   7
      TabsPerRow      =   3
      ActiveTab       =   1
      Tab(0).Subcontrol(0).Name=   "Text1"
      Tab(0).TotalSubcontrols=   1
      Tab(1).Caption  =   "Tab1"
      Tab(1).Subcontrol(0).Name=   "Text2"
      Tab(1).Subcontrol(0).Index=   "1"
      Tab(1).Subcontrol(1).Name=   "Text2"
      Tab(1).Subcontrol(1).Index=   "0"
      Tab(1).Subcontrol(2).Name=   "DataHWLabel1"
      Tab(1).TotalSubcontrols=   3
      Tab(2).Subcontrol(0).Name=   "Frame1"
      Tab(2).TotalSubcontrols=   1
      Tab(4).Caption  =   "Tab4"
      Tab(4).GoodState=   0   'False
      Tab(5).Caption  =   "Tab5"
      Tab(5).Enabled  =   0   'False
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   735
         Left            =   -74880
         TabIndex        =   14
         Top             =   1440
         Width           =   4575
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Text            =   "Text3"
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   1680
         Width           =   4575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   1320
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   -74880
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1680
         Width           =   4575
      End
      Begin ControlProject.DataHWLabel DataHWLabel1 
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         Caption         =   "DataHWLabel only on Tab1"
         BeginProperty DataFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label Visible On All Tabs"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   2280
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Text1"
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   1740
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Del"
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MyTab1"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Caption"
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   660
      Width           =   975
   End
End
Attribute VB_Name = "TestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

If LightTab1.TabCaption(1) = "By" Then
    LightTab1.TabCaption(1) = "VeryLongTabCaption1"
    LightTab1.TabCaption(2) = "VeryLongTabCaption2"
Else
    LightTab1.TabCaption(1) = "By"
    LightTab1.TabCaption(2) = "Michael K"
End If

End Sub

Private Sub Command2_Click()

LightTab1.ActiveTab = 1

End Sub

Private Sub Command3_Click()

If Command3.Caption = "Add" Then

    LightTab1.TotalTabs = LightTab1.TotalTabs + 1

    MsgBox LightTab1.TotalTabs

    Command3.Caption = "Enabled"

Else

    LightTab1.TabEnabled(1) = Not LightTab1.TabEnabled(1)

End If

End Sub

Private Sub Command4_Click()

If Command4.Caption = "Del" Then

    If LightTab1.TotalTabs > 1 Then LightTab1.TotalTabs = LightTab1.TotalTabs - 1

    MsgBox LightTab1.TotalTabs

    Command4.Caption = "GoodState"

Else

    LightTab1.TabGoodState(1) = Not LightTab1.TabGoodState(1)

End If

End Sub

Private Sub Command5_Click()

Select Case Command5.Caption
    Case "Text1"
        LightTab1.ActivateTabWithControl Text1.hWnd
        Command5.Caption = "Text2(0)"
    Case "Text2(0)"
        LightTab1.ActivateTabWithControl Text2(0).hWnd
        Command5.Caption = "Frame1"
    Case "Frame1"
        LightTab1.ActivateTabWithControl Frame1.hWnd
        Command5.Caption = "Text2(1)"
    Case "Text2(1)"
        LightTab1.ActivateTabWithControl Text2(1).hWnd
        Command5.Caption = "Text1"
End Select

End Sub

Private Sub Command6_Click()

Dim ChangeMode As Integer

ChangeMode = LightTab1.StateMode + 1

If ChangeMode = 4 Then ChangeMode = 0

LightTab1.StateMode = ChangeMode

TestForm.Caption = "LightTab - StateMode=" & LightTab1.StateMode & " - OffsetMode=" & LightTab1.OffsetMode & " - FlatMode=" & LightTab1.FlatMode

End Sub

Private Sub Command7_Click(ClickIndex As Integer)

LightTab1.TabVisible(ClickIndex) = Not LightTab1.TabVisible(ClickIndex)

End Sub

Private Sub Command8_Click()

Dim ChangeMode As Integer

ChangeMode = LightTab1.OffsetMode + 1

If ChangeMode = 3 Then ChangeMode = 0

LightTab1.OffsetMode = ChangeMode

TestForm.Caption = "LightTab - StateMode=" & LightTab1.StateMode & " - OffsetMode=" & LightTab1.OffsetMode & " - FlatMode=" & LightTab1.FlatMode

End Sub

Private Sub Command9_Click()

If LightTab1.TabsPerRow = 0 Then
    LightTab1.TabsPerRow = 3
Else
    LightTab1.TabsPerRow = 0
End If

End Sub

Private Sub Command10_Click()

LightTab1.InsideBorder = Not LightTab1.InsideBorder

End Sub

Private Sub Command11_Click()

Dim ChangeMode As Integer

ChangeMode = LightTab1.FlatMode + 1

If ChangeMode = 4 Then ChangeMode = 0

LightTab1.FlatMode = ChangeMode

TestForm.Caption = "LightTab - StateMode=" & LightTab1.StateMode & " - OffsetMode=" & LightTab1.OffsetMode & " - FlatMode=" & LightTab1.FlatMode

End Sub

Private Sub Form_Load()

TestForm.Caption = "LightTab - StateMode=" & LightTab1.StateMode & " - OffsetMode=" & LightTab1.OffsetMode & " - FlatMode=" & LightTab1.FlatMode

End Sub

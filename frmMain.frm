VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tooltip Class Test"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "R"
      Height          =   315
      Index           =   1
      Left            =   3360
      TabIndex        =   16
      Top             =   1980
      Width           =   315
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "R"
      Height          =   315
      Index           =   0
      Left            =   3360
      TabIndex        =   15
      Top             =   1620
      Width           =   315
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   4920
      TabIndex        =   14
      Top             =   3660
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2880
      Top             =   2820
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtTitle 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   12
      Text            =   "Enter title here"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CheckBox chkBalloon 
      Caption         =   "Balloon Style"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   11
      Top             =   3060
      Width           =   1305
   End
   Begin VB.PictureBox picTT 
      Height          =   2925
      Left            =   3720
      ScaleHeight     =   2865
      ScaleWidth      =   2385
      TabIndex        =   9
      Top             =   420
      Width           =   2445
   End
   Begin VB.CheckBox chkCenter 
      Caption         =   "Center"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   1305
   End
   Begin VB.CheckBox chkShowTitle 
      Caption         =   "Show Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   7
      Top             =   2460
      Width           =   1275
   End
   Begin VB.CommandButton cmdColour 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2220
      TabIndex        =   6
      Top             =   1980
      Width           =   1095
   End
   Begin VB.PictureBox picColour 
      BackColor       =   &H80000017&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1380
      ScaleHeight     =   255
      ScaleWidth      =   765
      TabIndex        =   4
      Top             =   1980
      Width           =   825
   End
   Begin VB.CommandButton cmdColour 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2220
      TabIndex        =   3
      Top             =   1620
      Width           =   1095
   End
   Begin VB.PictureBox picColour 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1380
      ScaleHeight     =   255
      ScaleWidth      =   765
      TabIndex        =   1
      Top             =   1620
      Width           =   825
   End
   Begin VB.TextBox txtTipText 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmMain.frx":0000
      Top             =   480
      Width           =   3375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   150
      X2              =   6210
      Y1              =   3510
      Y2              =   3510
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      X1              =   150
      X2              =   6210
      Y1              =   3495
      Y2              =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Tool Tip Text"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   13
      Top             =   180
      Width           =   2745
   End
   Begin VB.Label lblTest 
      Caption         =   "Test Here!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   10
      Top             =   60
      Width           =   2325
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Text Color:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Background Color:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   2
      Top             =   1650
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MyToolTip As CTooltip




'Form Load/Query unload
Private Sub Form_Load()
    Call Create
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set MyToolTip = Nothing 'Destroy tooltip object
End Sub

'Form wants to unload - will we let it??
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Are you sure you want to quit?", vbYesNo + vbDefaultButton2 + vbQuestion, "Quit?") = vbNo Then
        Cancel = 1  'Cancel the unload event
    End If
End Sub

'User clicked the quit button
Private Sub cmdQuit_Click()
    Unload Me   'Attempt to unload - catch in the Query unload event
End Sub

'User clicked the Show Title Tick box.
Private Sub chkShowTitle_Click()
    Me.txtTitle.Enabled = CBool(Me.chkShowTitle.Value)
    Call Create
End Sub
'The Title Tect box is changing. update the tooltip
Private Sub txtTitle_Change()
    Call Create
End Sub

'User wishes to choose a colour for thier tool tip.
Private Sub cmdColour_Click(Index As Integer)

    On Local Error GoTo cmdChangeBgColorClickError

        With CommonDialog1
            .CancelError = True
            .ShowColor
            picColour(Index).BackColor = .Color
        End With
        Call Create
        
cmdChangeBgColorClickExit:
    On Error Resume Next
    Exit Sub
cmdChangeBgColorClickError:
    Select Case Err.Number
        Case 32755
            Call MsgBox("User Canceled")
        Case Else
            Call MsgBox(Err.Number & ":" & Err.Description, vbCritical, "cmdChangeBgColor_Click")
    End Select
    Resume cmdChangeBgColorClickExit
End Sub

'Reset the tooltip colours to system defaults
Private Sub cmdReset_Click(Index As Integer)
    If Not MyToolTip Is Nothing Then
        If Index = 0 Then
            Me.picColour(Index).BackColor = MyToolTip.SystemToolTipBackColor
        Else
            Me.picColour(Index).BackColor = MyToolTip.SystemToolTipForeColor
        End If
    End If
    Call Create
End Sub


'Create the tool tip
Private Sub Create()
    
    On Error GoTo CreateError

    'Create a new tool tip object.
    Set MyToolTip = New CTooltip

    With MyToolTip
        'Set the Handle of the picture box for which we want the tooltip
        .HwndParentControl = picTT.hwnd
        
        'Set the text
        .Text = txtTipText.Text
        
        'Set the Tool Tip Type
        If CBool(chkBalloon.Value) Then
            .Style = TTBalloon
        Else
            .Style = TTStandard
        End If
        
        .Centered = CBool(chkCenter.Value)
        .BackColor = Me.picColour(0).BackColor  'Set the back colour
        .ForeColor = Me.picColour(1).BackColor  'Set the fore colour
        
        If CBool(chkShowTitle.Value) Then
            .Title = txtTitle.Text
        End If
        
        Call .Create    'Create the tooltip
    End With
CreateExit:
    On Error Resume Next
    Exit Sub
CreateError:
    Call MsgBox(Err.Number & ":" & Err.Description, vbCritical, "Create()")
    Resume CreateExit
End Sub

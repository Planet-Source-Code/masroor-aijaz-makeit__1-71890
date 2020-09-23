VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "MakeIt"
   ClientHeight    =   4770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   318
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   383
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3480
      TabIndex        =   10
      Top             =   0
      Width           =   735
   End
   Begin VB.Timer tmrDraw 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4320
      Top             =   240
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   4800
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Erase All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   9
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "Undo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   8
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox picColour 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1200
      ScaleHeight     =   285
      ScaleWidth      =   465
      TabIndex        =   7
      Top             =   0
      Width           =   495
   End
   Begin VB.ComboBox cboSize 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmGlass.frx":0000
      Left            =   360
      List            =   "frmGlass.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   0
      Width           =   735
   End
   Begin VB.PictureBox picPos 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   3600
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   5
      Top             =   0
      Width           =   225
      Begin VB.Line Line7 
         X1              =   4
         X2              =   11
         Y1              =   10
         Y2              =   10
      End
      Begin VB.Shape shpPos 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   225
         Index           =   2
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   225
      End
   End
   Begin VB.PictureBox picPos 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   3840
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   4
      Top             =   0
      Width           =   225
      Begin VB.Line Line11 
         X1              =   10
         X2              =   10
         Y1              =   4
         Y2              =   10
      End
      Begin VB.Line Line10 
         X1              =   4
         X2              =   4
         Y1              =   10
         Y2              =   4
      End
      Begin VB.Line Line9 
         X1              =   4
         X2              =   10
         Y1              =   4
         Y2              =   4
      End
      Begin VB.Line Line8 
         X1              =   10
         X2              =   4
         Y1              =   10
         Y2              =   10
      End
      Begin VB.Shape shpPos 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   225
         Index           =   1
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   225
      End
   End
   Begin VB.PictureBox picPos 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   4200
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   3
      Top             =   0
      Width           =   225
      Begin VB.Line Line12 
         X1              =   4
         X2              =   11
         Y1              =   10
         Y2              =   3
      End
      Begin VB.Line Line6 
         X1              =   4
         X2              =   11
         Y1              =   4
         Y2              =   11
      End
      Begin VB.Shape shpPos 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   225
         Index           =   0
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   225
      End
   End
   Begin VB.PictureBox picSize 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   4440
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   3960
      Width           =   225
      Begin VB.Line Line5 
         X1              =   45
         X2              =   180
         Y1              =   45
         Y2              =   180
      End
      Begin VB.Shape shpSize 
         BackColor       =   &H00D9D1A4&
         BackStyle       =   1  'Opaque
         Height          =   225
         Index           =   2
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   225
      End
   End
   Begin VB.PictureBox picSize 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   1
      Top             =   4080
      Width           =   225
      Begin VB.Line Line4 
         X1              =   7
         X2              =   7
         Y1              =   4
         Y2              =   11
      End
      Begin VB.Line Line3 
         X1              =   4
         X2              =   11
         Y1              =   7
         Y2              =   7
      End
      Begin VB.Shape shpSize 
         BackColor       =   &H00D9D1A4&
         BackStyle       =   1  'Opaque
         Height          =   225
         Index           =   3
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   225
      End
      Begin VB.Shape shpSize 
         BackStyle       =   1  'Opaque
         Height          =   225
         Index           =   1
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   225
      End
   End
   Begin VB.PictureBox picSize 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   0
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   0
      Width           =   225
      Begin VB.Line Line1 
         X1              =   4
         X2              =   11
         Y1              =   7
         Y2              =   7
      End
      Begin VB.Line Line2 
         X1              =   7
         X2              =   7
         Y1              =   4
         Y2              =   11
      End
      Begin VB.Shape shpSize 
         BackColor       =   &H00D9D1A4&
         BackStyle       =   1  'Opaque
         Height          =   225
         Index           =   0
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   225
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code is freeware and open source you can freely redistrubute it. But dont
'distrubute it with this name. Or if you have problem email me at masroora@paklog.co.cc


Option Explicit

Private WithEvents fEvents As Form
Attribute fEvents.VB_VarHelpID = -1
Private bNoResize As Boolean, bNoActivate As Boolean
Private bDragging As Boolean, sX As Single, sY As Single
Private ltbID As Long
Private sOldX As Single, sOldY As Single
Private picUndo As StdPicture

Private Sub cmdClear_Click()
    Me.Picture = Me.Image
    Set picUndo = Me.Picture
    Me.Cls
    Set Me.Picture = Nothing
End Sub

Private Sub cmdUndo_Click()
    Set Me.Picture = picUndo
End Sub


Private Sub cmdSave_Click()
    Dim sFile As String
    On Error GoTo ErrHandler
    With comDlg
        .DialogTitle = "Save As..."
        .InitDir = App.Path
        .DefaultExt = ".bmp"
        .Filter = "bmp|*.bmp"
        .ShowSave
        sFile = .FileName
    End With
    Me.Picture = Me.Image
    Set picUndo = Me.Picture
    SavePicture picUndo, sFile
ErrHandler:
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Picture = Me.Image
    Set picUndo = Me.Picture
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If sOldX = -1 Then
       sOldX = X
       sOldY = Y
    End If
    If Button = vbLeftButton Then Me.Line (X, Y)-(sOldX, sOldY), picColour.BackColor
    sOldX = X
    sOldY = Y
End Sub

Private Sub picColour_Click()
    With comDlg
        .DialogTitle = "Pick A Colour"
        .ShowColor
        picColour.BackColor = IIf(.Color = vbCyan, vbCyan + 1, .Color)
    End With
End Sub

Private Sub cboSize_Click()
    Me.DrawWidth = CInt(cboSize.List(cboSize.ListIndex))
End Sub

Private Sub picPos_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            If Me.WindowState = vbNormal Then Me.WindowState = vbMaximized Else Me.WindowState = vbNormal
        Case Else
            Me.WindowState = vbMinimized
    End Select
End Sub

Private Sub picSize_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 2 Then
        sX = X: sY = Y: bDragging = True
    Else
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
        Form_Resize
    End If
End Sub

Private Sub picSize_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const lMinWidth As Long = 300, lMinHeight As Long = 150
    Dim lWidth As Long, lHeight As Long
    If bDragging And Index = 2 Then
        lWidth = Me.Width + (X - sX)
        lHeight = Me.Height + (Y - sY)
        If lWidth < lMinWidth * Screen.TwipsPerPixelX Then lWidth = lMinWidth * Screen.TwipsPerPixelX
        If lHeight < lMinHeight * Screen.TwipsPerPixelX Then lHeight = lMinHeight * Screen.TwipsPerPixelX
        Me.Move Me.Left, Me.Top, lWidth, lHeight
    End If
End Sub

Private Sub picSize_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    bDragging = False
End Sub

Private Sub SetChildColour(ByVal lColour As Long)
    Dim N As Long, ctl As Control
    For N = picSize.LBound To picSize.UBound
        picSize(N).BackColor = lColour
    Next N
    For N = picPos.LBound To picPos.UBound
        picPos(N).BackColor = lColour
    Next N
    For Each ctl In Me.Controls
        If TypeOf ctl Is CommandButton Then ctl.BackColor = vbBlack
    Next ctl
End Sub

Private Sub PopulateControls()
    Dim N As Long
    sOldX = -1
    For N = 1 To 20
        cboSize.AddItem CStr(N)
    Next N
    cboSize.ListIndex = 4
    picColour.BackColor = vbBlack
    picColour.Height = cboSize.Height
    cmdUndo.Height = cboSize.Height
    cmdClear.Height = cboSize.Height
    cmdSave.Height = cboSize.Height
End Sub


Private Sub Form_Load()
    Set fEvents = frmBlank
    
    
    SetWindowLong Me.hWnd, GWL_STYLE, GetWindowLong(Me.hWnd, GWL_STYLE) Or WS_SYSMENU Or WS_MINIMIZEBOX
    
  
    Me.BackColor = vbCyan
    SetChildColour Me.BackColor
    
    SetTrans Me, , Me.BackColor
    SetTrans fEvents, 1
   

    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
    PopulateControls
        
    Me.Show
    ltbID = GetAppsID(Me.Caption)
    fEvents.Show
    Me.ZOrder
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Visible = False
    fEvents.Visible = False
    SetTrans Me
    SetTrans fEvents
    Unload fEvents
    Set fEvents = Nothing
    Set picUndo = Nothing
End Sub

Private Sub fEvents_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Picture = Me.Image
    Set picUndo = Me.Picture
End Sub

Private Sub fEvents_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If sOldX = -1 Then
       sOldX = X
       sOldY = Y
    End If
    If Not Button = vbLeftButton Then
        If IsOverCtl(Me, X, Y) Then Me.SetFocus
    Else
        Me.Line (X, Y)-(sOldX, sOldY), picColour.BackColor
    End If
    
    sOldX = X
    sOldY = Y
End Sub

Private Sub fEvents_Activate()
    SendMessage hTaskBar, TB_CHECKBUTTON, ltbID, 0&
    If Not GetNextWindow(fEvents.hWnd) = Me.hWnd Then Me.ZOrder
End Sub

Private Sub Form_Activate()
    If Not GetNextWindow(Me.hWnd) = fEvents.hWnd Then fEvents.ZOrder
End Sub

Private Sub fEvents_Resize()
    If bNoResize Then Exit Sub
    bNoResize = True
    With fEvents
        Me.Move .Left, .Top, .Width, .Height
    End With
    bNoResize = False
End Sub

Private Sub Form_Resize()
    Dim ctl As Control
    If bNoResize Then Exit Sub
    bNoResize = True
    With Me
        fEvents.Move .Left, .Top, .Width, .Height
    End With
    
    If Not Me.WindowState = vbMinimized Then

        picSize(1).Move 0, Me.ScaleHeight - picSize(1).Height
        picSize(2).Move Me.ScaleWidth - picSize(2).Width, Me.ScaleHeight - picSize(2).Height
        picPos(0).Move Me.ScaleWidth - picPos(0).Width, 0
        picPos(1).Move picPos(0).Left - picPos(1).Width - 4, 0
        picPos(2).Move picPos(1).Left - picPos(2).Width - 4, 0
        For Each ctl In Me.Controls
            If TypeOf ctl Is PictureBox Then ctl.Refresh
        Next ctl
        Me.Refresh
        picSize(0).Enabled = Not (Me.WindowState = vbMaximized)
        picSize(1).Enabled = Not (Me.WindowState = vbMaximized)
        picSize(2).Enabled = Not (Me.WindowState = vbMaximized)
    End If
    bNoResize = False
End Sub

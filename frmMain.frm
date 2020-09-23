VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Flash Player "
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8865
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7283.333
   ScaleMode       =   0  'User
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swf 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      _cx             =   7223
      _cy             =   8493
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   "000000"
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   4200
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17D2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":180D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":184CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18897
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18D2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":190C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   5040
      Top             =   1080
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1440
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":195B4
            Key             =   "Icon"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":199E3
            Key             =   "Iconz"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19E12
            Key             =   "Icons"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3960
      Top             =   2040
   End
   Begin MSComctlLib.Slider sld 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   4800
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   1
      Max             =   100
      TickStyle       =   3
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   2640
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31B4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37CD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":380D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E932
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":44ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":44E56
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   2  'Align Bottom
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   5955
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Play"
            Object.ToolTipText     =   "Play"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pause"
            Object.ToolTipText     =   "Pause"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Full"
            Object.ToolTipText     =   "Full Screen"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "URL"
            Object.ToolTipText     =   "Form The Web"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView List 
      Height          =   4695
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   8281
      View            =   3
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList2"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483625
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Object.Tag             =   "nn"
         Text            =   "Playlist"
         Object.Width           =   6165
      EndProperty
   End
   Begin VB.Label lbClick 
      Caption         =   "0"
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lb1 
      BackColor       =   &H80000006&
      Caption         =   "Current Frame:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Shape shp 
      BackColor       =   &H80000006&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   0
      Top             =   5400
      Width           =   8055
   End
   Begin VB.Menu mnFile 
      Caption         =   "File"
      Begin VB.Menu mnOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mn1 
         Caption         =   "-"
      End
      Begin VB.Menu mnExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnView 
      Caption         =   "View"
      Begin VB.Menu mnFull 
         Caption         =   "Full Screen"
      End
   End
   Begin VB.Menu mnAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NGUYEN QUANG HUNG
'87 NGO QUYEN ,THU DAU MOT,BINH DUONG
'CHUONG TRINH CHOI FLASH
Option Explicit
Dim Full As Boolean
Dim Start As Boolean
Dim Ready As Boolean
Dim swfH
Dim swfW
Dim swfT
Dim swfL
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_SYSMENU = &H80000
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    HWND_NOTOPMOST = -2
End Enum

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


        
Private Function ShowTitleBar(ByVal bState As Boolean)
    Dim lStyle As Long
    Dim tR As RECT
    GetWindowRect Me.hwnd, tR
    lStyle = GetWindowLong(Me.hwnd, GWL_STYLE)
    If (bState) Then
    Me.Caption = Me.Tag
    If Me.ControlBox Then
    lStyle = lStyle Or WS_SYSMENU
    End If
    If Me.MaxButton Then
    lStyle = lStyle Or WS_MAXIMIZEBOX
    End If
    If Me.MinButton Then
    lStyle = lStyle Or WS_MINIMIZEBOX
    End If
    If Me.Caption <> "" Then
    lStyle = lStyle Or WS_CAPTION
    End If
    Else
    Me.Tag = Me.Caption
    Me.Caption = ""
    lStyle = lStyle And Not WS_SYSMENU
    lStyle = lStyle And Not WS_MAXIMIZEBOX
    lStyle = lStyle And Not WS_MINIMIZEBOX
    lStyle = lStyle And Not WS_CAPTION
    End If
    SetWindowLong Me.hwnd, GWL_STYLE, lStyle
    SetWindowPos Me.hwnd, 0, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED
    Me.Refresh
    'Form_Resize
End Function
Private Sub Form_Load()
Full = False
Ready = False
Form_Resize
End Sub
Private Sub Form_Resize()
If Me.WindowState <> 1 Then
If Full = False Then
swfW = Me.ScaleWidth - swf.Left - List.Width - 100
swfH = Me.ScaleHeight - sld.Height - Toolbar.Height - shp.Height
swfT = 10
swfL = 10
Else
swfH = Me.ScaleHeight
swfW = Me.ScaleWidth
swfL = 0
swfT = 0
End If
swf.Move swfT, swfL, swfW, swfH
List.Top = swf.Top
List.Left = swf.Left + swf.Width + 10
List.Height = swf.Height
sld.Top = swf.Top + swf.Height + 10
sld.Width = Me.ScaleWidth - 10
shp.Top = sld.Top + sld.Height + 10
shp.Left = 10
shp.Width = Me.Width - 10
lb1.Height = shp.Height - 20
lb1.Top = shp.Top + 20
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Start = True Then
swf.Stop
End If
frmAbout.cmdOk.Value = True
End Sub

Private Sub List_DblClick()
If List.SelectedItem.Key <> "" Then
swf.Movie = List.SelectedItem.Key
'List.SelectedItem.SmallIcon = ImageList2.ListImages("Icons").Key
swf.Play
Start = True
Form1.Caption = "Flash Player - " & List.SelectedItem.text
End If
End Sub

Private Sub mnAbout_Click()
frmAbout.show
End Sub

Private Sub mnExit_Click()
swf.Stop
End
End Sub

Private Sub mnFull_Click()
Call FullScreen
End Sub

Private Sub mnOpen_Click()
lbClick = lbClick + 1
Dim Name As String
With dlg
.Flags = cdlOFNExplorer + cdlOFNAllowMultiselect
.Filter = "Shockware Flash File(*.swf)|*.swf|"
.MaxFileSize = 10000
.ShowOpen
End With
Name = StrConv(dlg.FileName, vbProperCase)
List.ListItems.Clear
If dlg.FileName = "" Then Exit Sub
     Dim a As Variant
     Dim P As Integer

     a = Split(dlg.FileName, vbNullChar)
     If UBound(a) = 0 Then
     
    List.ListItems.Add , dlg.FileName, dlg.FileTitle, "Icon", "Icon"
        Else
          For P = 1 To UBound(a)
               If Right(a(0), 1) <> "\" Then a(0) = a(0) & "\"
               List.ListItems.Add , a(0) + a(P), P & "   " & a(P), "Icon", "Icon"
          Next
     End If
     dlg.FileName = ""
     On Error Resume Next
    
End Sub
Private Sub sld_Scroll()
Dim x As Integer
x = (sld.Value * swf.TotalFrames) / 100
swf.GotoFrame (x)
swf.Play
Start = True
End Sub


Private Sub Timer1_Timer()
If Start = True Then
sld.Value = (swf.CurrentFrame * 100) / swf.TotalFrames
lb1.Caption = "Current Frame:" & swf.CurrentFrame & "/" & swf.TotalFrames
End If
End Sub

Private Sub Timer2_Timer()
If GetAsyncKeyState(vbKeyEscape) Then
ShowTitleBar True
Full = False
Me.WindowState = 0
mnFile.Visible = True
mnView.Visible = True
mnAbout.Visible = True
List.Visible = True
Toolbar.Visible = True
sld.Visible = True
shp.Visible = True
lb1.Visible = True
End If
End Sub



Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Play"
If lbClick <> 0 Then
swf.Movie = List.SelectedItem.Key
swf.Play
Else: Exit Sub
End If
Case "Stop"
If Start = True Then
swf.Stop
swf.GotoFrame (1)
sld.Value = (swf.CurrentFrame * 100) / swf.TotalFrames
End If
Start = False
Case "Full"
Call FullScreen
Case "Help"
Call mnAbout_Click
Case "URL"
Dim url
url = InputBox("Type the url", "Warning")
If url <> "" Then
List.ListItems.Add , url, url, "Icon", "Icon"
End If
Case "Pause"
swf.Stop
End Select
End Sub

Public Sub FullScreen()
ShowTitleBar False
Full = True
Me.WindowState = 2
Me.BorderStyle = 0
mnFile.Visible = False
mnView.Visible = False
mnAbout.Visible = False
List.Visible = False
Toolbar.Visible = False
sld.Visible = False
shp.Visible = False
lb1.Visible = False
End Sub




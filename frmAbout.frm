VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   360
      Top             =   1440
   End
   Begin VB.Frame Frame1 
      Caption         =   "About Author"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      Begin VB.Label lb1 
         Caption         =   "_"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   4815
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lb2 
      Caption         =   "0"
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
Unload Me
End Sub
Private Sub Timer1_Timer()
Dim text
Dim show
text = "NGUYEN QUANG HUNG + BINHDUONG + VIETNAM"
lb2 = lb2 + 1
If i <= Len(text) Then
show = Left(text, lb2)
End If
lb1.Caption = show & "_"
End Sub

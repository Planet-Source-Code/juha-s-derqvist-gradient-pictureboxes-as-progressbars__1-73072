VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "JUHAX.com - Progressbar Gradient"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture5 
      BackColor       =   &H80000007&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   1155
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   7395
      TabIndex        =   6
      Top             =   2160
      Width           =   7455
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   7395
      TabIndex        =   5
      Top             =   1320
      Width           =   7455
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   720
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   6345
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   1508
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Gradient Progress Bars (using a picturebox) including statusbars"
      Height          =   540
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7425
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4440
      Top             =   720
   End
   Begin VB.PictureBox Picture1 
      DrawStyle       =   5  'Transparent
      Height          =   255
      Left            =   4920
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4920
      ScaleHeight     =   255
      ScaleWidth      =   1575
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "gradient pictureboxes as progressbars - horizontal"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   11
      Top             =   5880
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "gradient vertical"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "exploding from inside"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "chainsaw"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
' (C) 2010 JUHAX.com
' progressbar in picturebox

Private Sub Command1_Click()
Picture1.Cls
Picture2.Cls
Picture3.Cls
Picture4.Cls
Picture5.Cls


' show pictureboxes
Picture1.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True

Command1.Enabled = False
ProgressBar1.Min = 0
ProgressBar1.Max = 100
ShowProgressInStatusBar True
Timer1.Enabled = True
End Sub

Private Sub ShowProgressInStatusBar(ByVal bShowProgressBar As Boolean)

    Dim tRC As RECT
    Dim tRC2 As RECT
    If bShowProgressBar Then
SendMessageAny StatusBar1.hwnd, SB_GETRECT, 0, tRC
SendMessageAny StatusBar1.hwnd, SB_GETRECT, 1, tRC2
With tRC
            .Top = (.Top * Screen.TwipsPerPixelY)
            .Left = (.Left * Screen.TwipsPerPixelX)
            .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
            .Right = (.Right * Screen.TwipsPerPixelX) - .Left
        End With
          With tRC2
            .Top = (.Top * Screen.TwipsPerPixelY)
            .Left = (.Left * Screen.TwipsPerPixelX)
            .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
            .Right = (.Right * Screen.TwipsPerPixelX) - .Left
        End With

        With Picture1
            SetParent .hwnd, StatusBar1.hwnd
            .Move tRC.Left, tRC.Top, tRC.Right, tRC.Bottom
            .Visible = True
        End With
        
          With Picture2
            SetParent .hwnd, StatusBar1.hwnd
            .Move tRC2.Left, tRC2.Top, tRC2.Right, tRC2.Bottom
            .Visible = True
        End With
    Else
        SetParent Picture1.hwnd, Me.hwnd
        SetParent Picture2.hwnd, Me.hwnd
        Picture1.Visible = False
        Picture2.Visible = False
        Picture3.Visible = False
        Picture4.Visible = False
        Picture5.Visible = False
        ProgressBar1.Visible = False
    End If
    
End Sub

 

 

Private Sub Form_Load()
' hide pictureboxes and progressbar
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
ProgressBar1.Visible = False
' color
Picture1.ForeColor = vbBlack
Picture1.CurrentX = Picture1.Width / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
ShowProgressInStatusBar False
End Sub

Private Sub Timer1_Timer()
    Static lCount As Long
    lCount = lCount + 1
    If lCount > 100 Then
        Timer1.Enabled = False
        ShowProgressInStatusBar False
        Command1.Enabled = True
        lCount = 0
    End If
    ' processbar picbox gradient here
    ProgressBar1.Value = lCount
   If Not lCount > 100 Then Gradient Me.Picture1, vbWhite, vbBlue, lCount, GradientHorizontal, True
   If Not lCount > 100 Then Gradient Me.Picture2, vbWhite, vbRed, lCount, GradientHorizontal, False

    If Not lCount > 100 Then Gradient2 Me.Picture3, vbWhite, vbBlue, lCount, GradientHorizontal, True
    If Not lCount > 100 Then Gradient3 Me.Picture4, vbWhite, vbRed, lCount, GradientHorizontal, True
    If Not lCount > 100 Then Gradient3 Me.Picture5, vbBlack, vbGreen, lCount, GradientVertical, True
 ' grey 12632256
End Sub


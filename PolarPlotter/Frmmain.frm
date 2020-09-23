VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pattern"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9900
   Icon            =   "Frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   " Drawing Option"
      Height          =   1170
      Left            =   3990
      TabIndex        =   7
      Top             =   8085
      Width           =   3795
      Begin VB.CheckBox Check3 
         Caption         =   "Clear pad for new drawing"
         Height          =   225
         Left            =   210
         TabIndex        =   17
         Top             =   840
         Value           =   1  'Checked
         Width           =   3480
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Show Axis"
         Height          =   225
         Left            =   1785
         TabIndex        =   16
         Top             =   525
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show Grid"
         Height          =   225
         Left            =   210
         TabIndex        =   15
         Top             =   525
         Value           =   1  'Checked
         Width           =   2010
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Options "
      Height          =   1170
      Left            =   105
      TabIndex        =   6
      Top             =   8085
      Width           =   3795
      Begin VB.TextBox txtx 
         Height          =   285
         Left            =   525
         TabIndex        =   12
         Text            =   "-4"
         Top             =   735
         Width           =   540
      End
      Begin VB.TextBox txtx2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1470
         TabIndex        =   11
         Text            =   "4"
         Top             =   735
         Width           =   540
      End
      Begin VB.TextBox txtt 
         Height          =   285
         Left            =   945
         TabIndex        =   9
         Text            =   "2"
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "x = "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   14
         Top             =   735
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "to"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1155
         TabIndex        =   13
         Top             =   735
         Width           =   165
      End
      Begin VB.Label Label3 
         Caption         =   "PI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1575
         TabIndex        =   10
         Top             =   315
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "t = 0 to "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   8
         Top             =   315
         Width           =   1065
      End
   End
   Begin VB.TextBox txtr 
      Height          =   330
      Left            =   630
      TabIndex        =   5
      Text            =   "2*Sin(3*t)"
      Top             =   7665
      Width           =   7155
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save As ..."
      Height          =   330
      Left            =   7875
      TabIndex        =   3
      Top             =   8925
      Width           =   1905
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   330
      Left            =   7875
      TabIndex        =   2
      Top             =   8505
      Width           =   1905
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Graph"
      Height          =   330
      Left            =   7875
      TabIndex        =   1
      Top             =   7665
      Width           =   1905
   End
   Begin VB.PictureBox Pad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7260
      Left            =   105
      ScaleHeight     =   7230
      ScaleWidth      =   9645
      TabIndex        =   0
      Top             =   105
      Width           =   9675
      Begin MSComDlg.CommonDialog cd1 
         Left            =   2940
         Top             =   1785
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   1890
         X2              =   4620
         Y1              =   3360
         Y2              =   4515
      End
   End
   Begin MSScriptControlCtl.ScriptControl SC1 
      Left            =   6090
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   120
      Left            =   105
      Top             =   7455
      Visible         =   0   'False
      Width           =   9675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "R(t) = "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   4
      Top             =   7770
      Width           =   540
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufileexit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuhelpabout 
         Caption         =   "About"
         Shortcut        =   ^{F1}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

'------------------- DIMS --------------------
Dim Zarib As Double

Private Sub Check1_Click()
    Call Command2_Click
End Sub

Private Sub Check2_Click()
    Pad.Cls
    Call Command2_Click
End Sub

Private Sub Command1_Click()
    pi = 4 * Atn(1)
    If Check3.Value = 1 Then
        Pad.Cls
        Call Command2_Click
    End If
    If Trim(txtr.Text) = "" Then
        MsgBox "You first must enter a function ! ", vbCritical, "Error"
        Exit Sub
    End If
    On Error GoTo errorhere
    txtr.Enabled = False
    Frame1.Enabled = False
    Frame2.Enabled = False
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Dim Zarib2 As Double
    Shape1.Visible = True
    Zarib2 = 9675 / Abs(Val(txtt.Text) * pi)
       
    Zarib = Pad.Width / Abs((Val(Trim(txtx.Text)) * 2))
    Shape1.Width = 0.01
    Dim xx, yy As Double
    Dim r As Double
    Dim Lastx, Lasty, Lastr As Double
    Dim t As Double
    
    t = 0.01
    SC1.ExecuteStatement ("t=" & t)
    Lastr = SC1.Eval(txtr.Text)
    Lastx = (Pad.Width / 2) + (Zarib * (Lastr * Cos(t)))
    Lasty = (Pad.Height / 2) - (Zarib * (Lastr * Sin(t)))
    For t = 0.01 To Val(txtt.Text) * pi Step 0.01
        SC1.ExecuteStatement ("t=" & t)
        r = SC1.Eval(txtr.Text)
        xx = (Pad.Width / 2) + (Zarib * (r * Cos(t)))
        yy = (Pad.Height / 2) - (Zarib * (r * Sin(t)))
        Pad.Line (Lastx, Lasty)-(xx, yy), vbBlue
        Lastx = xx
        Lasty = yy
        Shape1.Width = t * Zarib2
        DoEvents
        
    Next t
    txtr.Enabled = True
    Frame1.Enabled = True
    Frame2.Enabled = True
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Shape1.Visible = False
    Exit Sub
errorhere:
    MsgBox "An error has been occurred !" & vbCrLf & "Error number : " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbCritical, "Error"
    txtr.Enabled = True
    Frame1.Enabled = True
    Frame2.Enabled = True
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Shape1.Visible = False
End Sub

Private Sub Command2_Click()
    InitialPad
End Sub

Private Sub Command3_Click()
    On Error GoTo errorhere
    cd1.Filter = "Bitmap Image(*.bmp)|*.Bmp"
    cd1.ShowSave
    SavePicture Pad.Image, cd1.FileName
errorhere:
End Sub

Private Sub Form_Activate()
    InitialPad
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    On Error GoTo errhand
    Dim Apppath As String
    Apppath = App.Path
    If Left(Apppath, 1) <> 1 Then Apppath = Apppath & "\"
    Open Apppath & "init.txt" For Input As #1
        Dim str, strinit As String
        While Not EOF(1)
            Line Input #1, str
            strinit = strinit & vbCrLf & str
        Wend
     Close #1
    SC1.Language = "VBScript"
    SC1.AddCode (strinit)
    Exit Sub
errhand:
    MsgBox "An error has been occured!" & vbCrLf & "Error number : " & Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
End Sub

'InitialPad
Private Sub InitialPad()
    Pad.Cls
    Zarib = Pad.Width / Abs((Val(Trim(txtx.Text)) * 2))
    If Check1.Value = 1 Then
    Dim i As Double
    For i = Pad.Width / 2 To Pad.Width Step Zarib
        Pad.Line (i, 0)-(i, Pad.Width), RGB(230, 230, 230)
    Next i
    For i = Pad.Width / 2 To 0 Step -1 * Zarib
        Pad.Line (i, 0)-(i, Pad.Width), RGB(230, 230, 230)
    Next i
    For i = Pad.Height / 2 To Pad.Height Step Zarib
        Pad.Line (0, i)-(Pad.Width, i), RGB(230, 230, 230)
    Next i
    For i = Pad.Height / 2 To 0 Step -1 * Zarib
        Pad.Line (0, i)-(Pad.Width, i), RGB(230, 230, 230)
    Next
    End If
    If Check2.Value = 1 Then
    Pad.Line (Pad.Width / 2, 0)-(Pad.Width / 2, Pad.Height), RGB(180, 180, 180)
    Pad.Line (0, Pad.Height / 2)-(Pad.Width, Pad.Height / 2), RGB(180, 180, 180)
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnufileexit_Click()
    ans = MsgBox("Are you really want to exit?", vbQuestion + vbDefaultButton3 + vbYesNoCancel, "Exit")
    If ans = vbYes Then End
End Sub

Private Sub mnuhelpabout_Click()
    MsgBox "PolarPlotter" & vbCrLf & vbCrLf & "Programmer : Mahdi Eghbalahmadi" & vbCrLf & vbCrLf & "Air Univercity - Tehran-Iran", vbInformation, "About"
End Sub

Private Sub txtr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call Command1_Click
End Sub

Private Sub txtt_KeyPress(KeyAscii As Integer)
    validstr = "01234567890"
    If KeyAscii > 26 Then
        If InStr(validstr, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub

Private Sub txtt_LostFocus()
    If IsNumeric((txtt.Text)) Then
        Ignore = Ignore + 1
    Else
        MsgBox "You must enter a number for t ! ", vbCritical, "Error"
        txtt.SetFocus
        Exit Sub
    End If
  End Sub

Private Sub txtx_Change()
    If Trim(txtx.Text) <> "-" Then
        If Val(txtx.Text) >= 0 Then
            MsgBox "First number must be negative !", vbCritical, "Error"
            Exit Sub
        End If
    End If
    txtx2.Text = -1 * Val(txtx.Text)
    
End Sub

Private Sub txtx_KeyPress(KeyAscii As Integer)
    validstr = "01234567890.-+"
    If KeyAscii > 26 Then
        If InStr(validstr, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub

Private Sub txtx_LostFocus()
    If IsNumeric((txtx.Text)) Then
        Ignore = Ignore + 1
        Call Command2_Click
    Else
        MsgBox "You must enter a number for x ! ", vbCritical, "Error"
        txtx.SetFocus
    End If
End Sub

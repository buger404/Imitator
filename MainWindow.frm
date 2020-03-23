VERSION 5.00
Begin VB.Form MainWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ä£·Â"
   ClientHeight    =   6048
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   9804
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6048
   ScaleWidth      =   9804
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CheckBox Motion 
      Appearance      =   0  'Flat
      Caption         =   "ShowMotion"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5280
      TabIndex        =   9
      Top             =   5520
      Width           =   1596
   End
   Begin VB.CheckBox Masked 
      Appearance      =   0  'Flat
      Caption         =   "ShowMask"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5280
      TabIndex        =   8
      Top             =   5136
      Width           =   1596
   End
   Begin VB.Timer UpdateTimer 
      Enabled         =   0   'False
      Interval        =   16
      Left            =   4560
      Top             =   168
   End
   Begin VB.PictureBox Outputs 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawWidth       =   5
      ForeColor       =   &H80000008&
      Height          =   3636
      Left            =   5280
      ScaleHeight     =   3636
      ScaleWidth      =   3876
      TabIndex        =   1
      Top             =   1368
      Width           =   3876
   End
   Begin VB.PictureBox Inputs 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawWidth       =   5
      ForeColor       =   &H80000008&
      Height          =   3636
      Left            =   504
      ScaleHeight     =   3636
      ScaleWidth      =   3876
      TabIndex        =   0
      Top             =   1392
      Width           =   3876
   End
   Begin VB.Label ImiBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF3E64&
      Caption         =   "IMITATE"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   7488
      TabIndex        =   7
      Top             =   5352
      Width           =   1692
   End
   Begin VB.Label StateText 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEDB1A&
      BackStyle       =   0  'Transparent
      Caption         =   "0 layers"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   528
      TabIndex        =   6
      Top             =   192
      Width           =   8628
   End
   Begin VB.Label GoBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF3E64&
      Caption         =   "ACTIVE"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   7440
      TabIndex        =   5
      Top             =   648
      Width           =   1692
   End
   Begin VB.Label LoadBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF3E64&
      Caption         =   "LOAD"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5280
      TabIndex        =   4
      Top             =   648
      Width           =   1812
   End
   Begin VB.Label SaveBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CEDB1A&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2664
      TabIndex        =   3
      Top             =   648
      Width           =   1692
   End
   Begin VB.Label ClearBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CEDB1A&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   528
      TabIndex        =   2
      Top             =   648
      Width           =   1812
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Dot
    X As Single
    Y As Single
End Type
Private Type Dots
    P() As Dot
End Type
Private Type DotsFile
    G() As Dots
End Type
Private Type ImiFile
    Dfs() As DotsFile
End Type
Dim Eg As DotsFile
Dim DF As ImiFile
Dim FileName As String
Private Sub ClearBtn_Click()
    ReDim Eg.G(0)
    ReDim Eg.G(0).P(0)
    Inputs.Cls
End Sub

Private Sub Form_Load()
    ReDim Eg.G(0)
    ReDim Eg.G(0).P(0)
    ReDim DF.Dfs(0)
End Sub

Public Sub AddDot(X As Single, Y As Single)
    Dim GI As Integer
    GI = UBound(Eg.G)
    ReDim Preserve Eg.G(GI).P(UBound(Eg.G(GI).P) + 1)
    With Eg.G(GI).P(UBound(Eg.G(GI).P))
        .X = X
        .Y = Y
    End With
    If UBound(Eg.G(GI).P) > 1 Then
        Inputs.Line (Eg.G(GI).P(UBound(Eg.G(GI).P) - 1).X, Eg.G(GI).P(UBound(Eg.G(GI).P) - 1).Y)- _
                    (Eg.G(GI).P(UBound(Eg.G(GI).P)).X, Eg.G(GI).P(UBound(Eg.G(GI).P)).Y)
    End If
End Sub

Private Sub GoBtn_Click()
    If UBound(DF.Dfs) < 1 Then MsgBox "No layers !", 16: Exit Sub
    
    UpdateTimer.Enabled = Not UpdateTimer.Enabled
    GoBtn.BackColor = IIf(UpdateTimer, RGB(128, 128, 128), LoadBtn.BackColor)
End Sub

Private Sub ImiBtn_Click()
    If UBound(DF.Dfs) < 1 Then MsgBox "No layers !", 16: Exit Sub
    Outputs.ForeColor = RGB(128, 128, 128)
    Call Imitate
End Sub

Private Sub Inputs_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyS Then Call SaveBtn_Click
End Sub

Private Sub Inputs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    AddDot X, Y
End Sub

Private Sub Inputs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then AddDot X, Y
End Sub

Private Sub Inputs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    AddDot X, Y
    ReDim Preserve Eg.G(UBound(Eg.G) + 1)
    ReDim Eg.G(UBound(Eg.G)).P(0)
End Sub

Private Sub LoadBtn_Click()
    Dim file As String
    file = InputBox("your handwritting name")
    FileName = file
    If Dir(App.Path & "\" & file & ".data") = "" Then
        MsgBox "An empty handwritting , it'll now be created .", 48
        ReDim Eg.G(0)
        ReDim Eg.G(0).P(0)
        ReDim DF.Dfs(0)
        StateText.Caption = UBound(DF.Dfs) & " layers"
    Else
        Open App.Path & "\" & file & ".data" For Binary As #1
        Get #1, , DF
        Close #1
        StateText.Caption = UBound(DF.Dfs) & " layers"
        MsgBox "Complete", 64
    End If
End Sub

Private Sub SaveBtn_Click()
    If FileName = "" Then MsgBox "Load a new handwritting first.", 16: Exit Sub

    If UBound(DF.Dfs) > 1 Then
        If UBound(Eg.G) <> UBound(DF.Dfs(UBound(DF.Dfs) - 1).G) Then MsgBox "Unmatched line counts !", 16: Exit Sub
    End If
    
    ReDim Preserve DF.Dfs(UBound(DF.Dfs) + 1)
    DF.Dfs(UBound(DF.Dfs)) = Eg
    StateText.Caption = UBound(DF.Dfs) & " layers"
    Open App.Path & "\" & FileName & ".data" For Binary As #1
    Put #1, , DF
    Close #1
    
    Call ClearBtn_Click
End Sub

Private Sub Imitate()
    Outputs.Cls
    
    Dim m As DotsFile, t As Dots, d As Dot, m2 As DotsFile
    m = DF.Dfs(Int(Rnd * UBound(DF.Dfs)) + 1)
    Dim ox As Long, oy As Long, depth As Single, SS As Long, pitch As Long
    SS = Outputs.Width * Outputs.Height / 1500
    
    If Masked.Value Then m2 = m
    
    Randomize
    For i = 0 To UBound(m.G) - 1
        pitch = UBound(m.G(i).P) / 10
        For s = 1 To UBound(m.G(i).P)
            For j = 1 To UBound(DF.Dfs)
                t = DF.Dfs(j).G(i)
                d = t.P(s / UBound(m.G(i).P) * UBound(t.P))
                With m.G(i).P(s)
                    ox = .X: oy = .Y
                    .X = .X + (d.X - .X) * Rnd * 0.233
                    .Y = .Y + (d.Y - .Y) * Rnd * 0.233
                    If Motion.Value And s Mod pitch = 0 Then
                        Outputs.DrawWidth = 1
                        depth = Abs(.X - ox) * Abs(.Y - oy) / SS
                        If depth > 1 Then depth = 1
                        Outputs.Line (ox, oy)-(.X, .Y), RGB(255 * depth, 0, 0)
                        ox = .X: oy = .Y
                        Outputs.DrawWidth = 5
                    End If
                    If s > 1 Then
                        .X = .X + (m.G(i).P(s - 1).X - .X) * 0.45
                        .Y = .Y + (m.G(i).P(s - 1).Y - .Y) * 0.45
                    End If
                    If Motion.Value And s Mod pitch = 0 Then
                        Outputs.DrawWidth = 1
                        depth = Abs(.X - ox) * Abs(.Y - oy) / SS
                        If depth > 1 Then depth = 1
                        Outputs.Line (ox, oy)-(.X, .Y), RGB(0, 255 * depth, 0)
                        Outputs.DrawWidth = 5
                    End If
                End With
            Next
        Next
    Next
    
    'Outputs
    If Masked.Value Then
        Inputs.Cls
        Call OutputDots(m2, Inputs)
        Outputs.ForeColor = RGB(230, 230, 230)
        Call OutputDots(m2, Outputs)
    End If
    
    Outputs.ForeColor = LoadBtn.BackColor
    Call OutputDots(m, Outputs)
End Sub
Private Sub OutputDots(m As DotsFile, Pad As PictureBox)
    For i = 0 To UBound(m.G) - 1
        For s = 2 To UBound(m.G(i).P)
            Pad.Line (m.G(i).P(s - 1).X, m.G(i).P(s - 1).Y)-(m.G(i).P(s).X, m.G(i).P(s).Y)
        Next
    Next
End Sub
Private Sub UpdateTimer_Timer()
    Call Imitate
End Sub

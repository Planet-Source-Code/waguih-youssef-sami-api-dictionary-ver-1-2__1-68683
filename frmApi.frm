VERSION 5.00
Begin VB.Form frmApi 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C000&
   Caption         =   "Api Tutorial"
   ClientHeight    =   6510
   ClientLeft      =   -1710
   ClientTop       =   -825
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   613
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   1080
      TabIndex        =   23
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   4680
      Width           =   5895
   End
   Begin VB.VScrollBar VS 
      Height          =   6375
      LargeChange     =   525
      Left            =   8880
      TabIndex        =   21
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   16
      Left            =   150
      TabIndex        =   20
      Top             =   5940
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      TabIndex        =   19
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H80000005&
      Height          =   1800
      Index           =   15
      Left            =   1035
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   5130
      Width           =   7680
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   2865
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   4230
      Width           =   5895
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   1080
      TabIndex        =   11
      Top             =   4230
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   3765
      Width           =   5895
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   1080
      TabIndex        =   9
      Top             =   3750
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   3300
      Width           =   5895
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1080
      TabIndex        =   7
      Top             =   3300
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2820
      Width           =   5895
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1080
      TabIndex        =   5
      Top             =   2820
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2865
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2385
      Width           =   5895
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1080
      TabIndex        =   3
      Top             =   2370
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   7695
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      ForeColor       =   &H000040C0&
      Height          =   375
      Index           =   0
      Left            =   1785
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2415
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   1125
      Width           =   7635
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "EXAMPLE"
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PARAM:"
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Syntax"
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Purpose"
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Function Name"
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAddNew 
         Caption         =   "Add New"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuUpDate 
         Caption         =   "Update"
      End
   End
End
Attribute VB_Name = "frmApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset

Dim VSLVal As Integer 'Last Value of Vertical Scrollbar
Dim VMax As Integer 'Max Value for Vertical Scrollbar

Public Sub See()
For i = 0 To 15
Text(i).Visible = True
Next

For i = 0 To 15
If rs(i) <> "" Then
Text(i).Text = rs(i)
End If
Next i



Label5.Top = Text(15).Top

If Dir(App.Path & "\" & rs(0) & ".jpg") <> "" Then
Image1.Picture = LoadPicture(App.Path & "\" & rs(0) & ".jpg")
Else
Image1.Picture = LoadPicture("")
End If
End Sub

Public Sub Fill_Combo()
Dim MSQL As String
MySQL = "SELECT Name from Functions ORDER By Name"
Set rs1 = db.OpenRecordset(MySQL)
rs1.MoveFirst
Do While Not rs1.EOF
Combo1.AddItem rs1("Name")
rs1.MoveNext
Loop

End Sub

Private Sub Combo1_Click()
Dim MSQL As String
MySQL = "SELECT * from Functions WHERE Name = " & Chr(34) & Combo1.Text & Chr(34)
Set rs = db.OpenRecordset(MySQL)
sClear
See
 For i = 4 To 14 Step 2 'For i = 3 To 14
NoOfLines (i)
Next i
Vanish
Label5.Top = Text(15).Top
End Sub

Private Sub Form_Load()
mnuUpDate.Enabled = False
Set db = OpenDatabase(App.Path & "\API.mdb")
Set rs = db.OpenRecordset("Functions", dbOpenTable)
rs.MoveFirst
sClear
See
Vanish
Fill_Combo

'******************************
On Error Resume Next
Dim Cntrl As Control

For Each Cntrl In Me.Controls
 If (Not TypeOf Cntrl Is VScrollBar) Then
  If (Cntrl.Top + Cntrl.Height + 125) > VMax Then
   VMax = Cntrl.Top + Cntrl.Height + 125
  End If
  
 End If
Next
On Error GoTo 0

'Set the Default Values for ScrollBar

With VS
.Height = Me.ScaleHeight
.Left = Me.ScaleWidth - .Width
.Top = 0
.Max = VMax - .Height
If .Max < 0 Then .Max = 0
End With

'If Max Value <= 0 then Disable ScrollBar
If VS.Max <= 0 Then VS.Enabled = False

End Sub

Public Sub sClear()
For i = 0 To 15
Text(i).Visible = True
Text(i).Text = ""
Next i
End Sub

Public Sub Vanish()
For i = 3 To 14
If Text(i).Text = "" Then
Text(i).Visible = False
End If
Next i
ScaleMode = 3
Label5.Top = Text(15).Top
End Sub

Public Sub NoOfLines(x As Integer)
Dim a
a = 1
For r = 1 To Len(Text(x).Text) - 1
If Mid(Text(x).Text, r, 2) = vbCrLf Then
a = a + 1
End If
Next r
If a > 1 Then
Text(x).Height = 25 * a
Else
Text(x).Height = 25
End If
Text(x + 1).Top = Text(x).Top + Text(x).Height + 2
Text(x + 2).Top = Text(x).Top + Text(x).Height + 2
Text(x - 1).Top = Text(x).Top
End Sub

Private Sub VSPositionChanged()
On Error Resume Next
Dim Val As Integer, Cntrl As Control
Val = VSLVal - VS.Value
For Each Cntrl In Me.Controls
 If TypeOf Cntrl Is VScrollBar Then
 'Do Nothing
 ElseIf TypeOf Cntrl Is Line Then
  Cntrl.Y1 = Cntrl.Y1 + Val
  Cntrl.Y2 = Cntrl.Y2 + Val
 Else
  Cntrl.Top = Cntrl.Top + Val
 End If
Next
VSLVal = VS.Value
End Sub

Private Sub mnuAddNew_Click()
mnuEdit.Enabled = False
mnuUpDate.Enabled = True
sClear
For i = 0 To 15
Text(i).Visible = True
Next
Image1.Picture = LoadPicture("")
rs.AddNew
End Sub

Private Sub mnuEdit_Click()
mnuAddNew.Enabled = False
mnuUpDate.Enabled = True
rs.Edit
End Sub

Private Sub mnuUpDate_Click()
mnuEdit.Enabled = True
mnuAddNew.Enabled = True

On Error GoTo errHandler
For i = 0 To 15
If Text(i) <> "" Then
rs(i) = Text(i)
End If
Next i
rs.Update
Fill_Combo
Combo1.Refresh
mnuUpDate.Enabled = False
Exit Sub

errHandler:
MsgBox "you should press Edit Or Add New First", vbOKOnly
Exit Sub
End Sub

Private Sub VS_Change()
VSPositionChanged
Text(0).SetFocus
End Sub

Private Sub VS_Scroll()
VSPositionChanged
Text(0).SetFocus
End Sub
Private Sub Form_Resize()
If Me.WindowState = 1 Then
 Exit Sub 'Exit if Minimized
ElseIf Me.Height < 2525 Then
 Me.Height = 2525 'Default Height
ElseIf Me.Width < 2525 Then
 Me.Width = 2525 'Default Width
End If


With VS
.Height = Me.ScaleHeight
.Left = Me.ScaleWidth - .Width
.Top = 0
.Max = VMax - .Height
End With

'If Max Value <= 0 then Disable ScrollBar
If VS.Max <= 0 Then
 VS.Enabled = False
 VS.Max = 0
Else
 VS.Enabled = True
End If
End Sub


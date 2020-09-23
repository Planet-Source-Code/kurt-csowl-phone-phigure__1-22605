VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HSU Phone Phigure"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4170
   Icon            =   "PhonePhigure.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdb 
      Left            =   1845
      Top             =   1335
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrColor 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1485
      Top             =   1335
   End
   Begin VB.TextBox txtData 
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1260
      Width           =   3900
   End
   Begin VB.CheckBox chkFile 
      Caption         =   "Write File"
      Height          =   540
      Left            =   2925
      TabIndex        =   3
      Top             =   180
      Width           =   1020
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Phigure"
      Default         =   -1  'True
      Height          =   495
      Left            =   1950
      TabIndex        =   2
      Top             =   225
      Width           =   810
   End
   Begin VB.Frame Frame1 
      Caption         =   "Phone Number"
      Height          =   645
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1770
      Begin VB.TextBox txtNumber 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   135
         TabIndex        =   1
         Text            =   "5553825"
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.Label lblWeb 
      Alignment       =   2  'Center
      Caption         =   "hsunderground.box.sk"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   6
      Top             =   4515
      Width           =   3270
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Ready"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1575
      TabIndex        =   4
      Top             =   855
      Width           =   1665
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalc_Click()
tmrColor.Enabled = True
Dim DaFile As String
If Len(txtNumber.Text) > 7 Or chkFile.Value = 1 Then
If Len(txtNumber.Text) > 7 Then MsgBox "Because the number is greater than 7 digits it will necessitate a file"
cdb.DialogTitle = "File to Create"
cdb.Filter = "Text File|*.txt|All Files|*.*"
cdb.ShowSave
DaFile = cdb.FileName
If DaFile = "" Then Exit Sub
Open DaFile For Output As #1
Dim Placer As String, Tempr As String
Dim Holderr As Integer, Possibler As Long
Dim PlacerLen As Integer, iCounterr As Long

'Create all 1's
Placer = String(Len(txtNumber), "1")
'initialize using all 1's
Print #1, ComAndSet(Placer)
iCounterr = 1
PlacerLen = Len(Placer)
Possibler = 3 ^ PlacerLen
Me.MousePointer = 11
Do
    'get last Placer value
    Holderr = Val(Right(Placer, 1))
    'increment by 1
    Holderr = Holderr + 1
    'Use mid STATEMENT to rePlacer it back
    Mid(Placer, PlacerLen, 1) = Holderr
    
    'Check for 4's and then move down to a 1 and increment next
    For x = PlacerLen To 1 Step -1
        Tempr = Mid(Placer, x, 1)
        If Tempr = "4" Then
            If x = 1 Then
                lblStatus.Caption = "DONE!"
                Close
                Me.MousePointer = 0
                Exit Sub
            End If
            Holderr = Val(Mid(Placer, x - 1, 1))
            Holderr = Holderr + 1
            Mid(Placer, x, 1) = "1"
            Mid(Placer, x - 1, 1) = CStr(Holderr)
        End If
    Next x
    Tempr = ComAndSet(Placer)
    Print #1, Tempr
    iCounterr = iCounterr + 1
    lblStatus = Round((iCounterr / Possibler) * 100, 1) & "%"
    lblStatus.Refresh
Loop
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' If not by file then use text box, SAME CODE
Else
Dim Place As String, Temp As String
Dim Holder As Integer, Possible As Long
Dim PlaceLen As Integer, iCounter As Long

'Create all 1's
Place = String(Len(txtNumber), "1")
'initialize using all 1's
txtData.Text = ComAndSet(Place) & vbCrLf
iCounter = 1
PlaceLen = Len(Place)
Possible = 3 ^ PlaceLen
Me.MousePointer = 11
Do
    'get last place value
    Holder = Val(Right(Place, 1))
    'increment by 1
    Holder = Holder + 1
    'Use mid STATEMENT to replace it back
    Mid(Place, PlaceLen, 1) = Holder
    
    'Check for 4's and then move down to a 1 and increment next
    For x = PlaceLen To 1 Step -1
        Temp = Mid(Place, x, 1)
        If Temp = "4" Then
            If x = 1 Then
                lblStatus.Caption = "DONE!"
                Me.MousePointer = 0
                Exit Sub
            End If
            Holder = Val(Mid(Place, x - 1, 1))
            Holder = Holder + 1
            Mid(Place, x, 1) = "1"
            Mid(Place, x - 1, 1) = CStr(Holder)
        End If
    Next x
    Temp = ComAndSet(Place)
    txtData.Text = txtData.Text & Temp & vbCrLf
    iCounter = iCounter + 1
    lblStatus = Round((iCounter / Possible) * 100, 1) & "%"
    lblStatus.Refresh
Loop
'''''''''''''''''''' 3 ^ placeLen = possibles
End If
End Sub

Private Sub lblWeb_Click()
Shell "Explorer http://hsunderground.box.sk", vbNormalFocus
End Sub

Private Sub tmrColor_Timer()
Static x As Integer
x = x Mod 15 + 1
lblWeb.BackColor = QBColor(x)
lblWeb.Refresh
End Sub

Private Sub txtNumber_Change()
Dim Temp As String
For x = 1 To Len(txtNumber.Text)
    Temp = Mid(txtNumber.Text, x, 1)
    If Not IsNumeric(Temp) Then txtNumber.Text = Replace(txtNumber.Text, Temp, "")
Next x
End Sub

Private Function ComAndSet(Numberz As String) As String
'Convert Numeric Code into Letterz
For x = 1 To Len(txtNumber.Text)
    ComAndSet = ComAndSet & GetLet(Mid(txtNumber.Text, x, 1), Mid(Numberz, x, 1))
Next x

End Function

Private Function GetLet(Keynum As Integer, Placenum As Integer) As String
' This function is fed the Keypad value and which Letter to use  and it returns the answer
Select Case Keynum
Case 0:
    GetLet = 0
Case 1:
    GetLet = 1
Case 2:
    If Placenum = 1 Then
        GetLet = "A"
    ElseIf Placenum = 2 Then
        GetLet = "B"
    Else
        GetLet = "C"
    End If
Case 3:
    If Placenum = 1 Then
        GetLet = "D"
    ElseIf Placenum = 2 Then
        GetLet = "E"
    Else
        GetLet = "F"
    End If
Case 4:
    If Placenum = 1 Then
        GetLet = "G"
    ElseIf Placenum = 2 Then
        GetLet = "H"
    Else
        GetLet = "I"
    End If
Case 5:
    If Placenum = 1 Then
        GetLet = "J"
    ElseIf Placenum = 2 Then
        GetLet = "K"
    Else
        GetLet = "L"
    End If
Case 6:
    If Placenum = 1 Then
        GetLet = "M"
    ElseIf Placenum = 2 Then
        GetLet = "N"
    Else
        GetLet = "O"
    End If
Case 7:
    If Placenum = 1 Then
        GetLet = "P"
    ElseIf Placenum = 2 Then
        GetLet = "R"
    Else
        GetLet = "S"
    End If
Case 8:
    If Placenum = 1 Then
        GetLet = "T"
    ElseIf Placenum = 2 Then
        GetLet = "U"
    Else
        GetLet = "V"
    End If
Case 9:
    If Placenum = 1 Then
        GetLet = "W"
    ElseIf Placenum = 2 Then
        GetLet = "X"
    Else
        GetLet = "Y"
    End If
End Select
End Function


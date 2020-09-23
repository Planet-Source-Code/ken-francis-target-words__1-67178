VERSION 5.00
Begin VB.Form frmTargetWords 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Target Words"
   ClientHeight    =   3975
   ClientLeft      =   5265
   ClientTop       =   3825
   ClientWidth     =   5370
   Icon            =   "frmTargetWords.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAnagram 
      Caption         =   "Single word Anagrams"
      Height          =   315
      Left            =   2340
      TabIndex        =   7
      Top             =   1860
      Width           =   2025
   End
   Begin VB.TextBox txtCommon2 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5010
      TabIndex        =   4
      Top             =   930
      Width           =   225
   End
   Begin VB.TextBox txtCommon1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4740
      TabIndex        =   3
      Top             =   930
      Width           =   225
   End
   Begin VB.TextBox txtMinimum 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4740
      TabIndex        =   6
      Top             =   1380
      Width           =   315
   End
   Begin VB.CheckBox chkMinimum 
      Caption         =   "Minimum letters per word"
      Height          =   375
      Left            =   2340
      TabIndex        =   5
      Top             =   1365
      Width           =   2100
   End
   Begin VB.CheckBox chkCommon 
      Caption         =   "Each word to contain letter/s"
      Height          =   330
      Left            =   2340
      TabIndex        =   2
      Top             =   915
      Width           =   2400
   End
   Begin VB.CheckBox chkOnce 
      Caption         =   "Use each letter only once per word"
      Height          =   330
      Left            =   2340
      TabIndex        =   1
      Top             =   465
      Width           =   2865
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   360
      Left            =   3270
      TabIndex        =   9
      Top             =   2955
      Width           =   1260
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   345
      Left            =   3270
      TabIndex        =   10
      Top             =   3495
      Width           =   1275
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Words"
      Height          =   345
      Left            =   3270
      TabIndex        =   8
      Top             =   2400
      Width           =   1260
   End
   Begin VB.ListBox lstResults 
      Height          =   2790
      Left            =   135
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   825
      Width           =   2025
   End
   Begin VB.TextBox txtEnter 
      Height          =   270
      Left            =   150
      TabIndex        =   0
      Top             =   495
      Width           =   2010
   End
   Begin VB.Label lblResults 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   150
      TabIndex        =   13
      Top             =   3690
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter target word or letters"
      Height          =   195
      Left            =   165
      TabIndex        =   11
      Top             =   195
      Width           =   1860
   End
End
Attribute VB_Name = "frmTargetWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Make words and anagrams from target word or letter string
'Ken Francis 2006. email: windmill_end@hotmail.com
Option Explicit

Private Sub chkAnagram_Click()
Dim TempStr As String
TempStr = txtEnter.Text

If chkAnagram.Value Then
    If txtEnter = "" Then
    MsgBox "Enter Target Word First", vbExclamation, "Target Words"
    txtEnter.SetFocus
    chkAnagram.Value = 0
    Exit Sub
    Else
    txtEnter.Enabled = False
    lstResults.Clear
    cmdFind.Caption = "Find Anagrams"
End If
    
    With chkOnce
        .Value = 1
        .Enabled = False
    End With
    With chkCommon
        .Value = 0
        .Enabled = False
    End With
    With txtCommon1
        .Text = ""
        .Enabled = False
    End With
    With txtCommon2
        .Text = ""
        .Enabled = False
    End With
    With chkMinimum
        .Value = 1
        .Enabled = False
    End With
    With txtMinimum
        .Text = Len(txtEnter)
        .Enabled = False
    End With
Else
    cmdFind.Caption = "Find Words"
    ClearAll
    txtEnter.Text = TempStr         'restore target word after clear(optional)
End If

End Sub

Private Sub ChkCommon_Click()

If chkCommon.Value Then
    With txtCommon1
        .Enabled = True
        .Text = ""
        .SetFocus
    End With
        With txtCommon2
        .Enabled = True
        .Text = ""
    End With
Else
    With txtCommon1
        .Text = ""
        .Enabled = False
    End With
    With txtCommon2
        .Text = ""
        .Enabled = False
    End With
End If

End Sub

Private Sub chkMinimum_Click()
If chkMinimum.Value Then
    With txtMinimum
        .Enabled = True
        .SetFocus
    End With
Else
     With txtMinimum
      .Text = ""
      .Enabled = False
     End With
End If
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdClear_Click()
ClearAll
End Sub
Private Sub cmdFind_Click()
Dim ListWord As String, EnteredWord As String
Dim CommonLetter1 As String, LenMinimum As Integer
Dim CommonLetter2 As String, StartTime As Single
Dim MinimumLength As Integer, LenSampleWord As Integer
Dim SampleWord As String, ChkStr As String
Dim k As Integer, Counter As Long, ModeStr As String

On Error GoTo ErrHandler

If Val(txtMinimum.Text) > Len(txtEnter.Text) Then
    MsgBox "When 'Use each letter only once' is selected" & vbLf & vbCr & _
    "words cannot have more letters than Target Word", vbInformation, "Target Words"
    With txtMinimum
        .Text = ""
        .SetFocus
    End With
    Exit Sub
End If

If txtEnter = "" Then
    MsgBox "Target Word must be Entered!", vbInformation, "Target Words"
    txtEnter.SetFocus
    Exit Sub
End If

Me.MousePointer = vbHourglass
StartTime = Timer
lblResults.Caption = ""
lstResults.Clear
CommonLetter1 = LCase(txtCommon1.Text)
CommonLetter2 = LCase(txtCommon2.Text)
SampleWord = LCase(Trim(txtEnter.Text))
LenMinimum = Val(txtMinimum.Text)
LenSampleWord = Len(SampleWord)
EnteredWord = SampleWord

Open App.Path & "\wordsdic.dic" For Input As #1

Do Until EOF(1)

    Input #1, ListWord
     
        Do
        
'*******************************************************************************************
'*******************************************************************************************
            'This word search engine uses the Sherlock Holmes principle "Eliminate all other
            'factors, and the one which remains must be the truth."
            '"VB5 users will need to add code in a BAS module to emulate the VB6 'Replace'
            'function. See below.
        
            'Eliminate words with less than minimum or more than maximum letters
            If Len(ListWord) < LenMinimum Then Exit Do
            If chkOnce.Value And Len(ListWord) > LenSampleWord Then Exit Do

            'Eliminate words with no common letter/s(if selected).
            'The faster "vbBinaryCompare" is used in "Instr" & "Replace" functions
            If chkCommon.Value Then
               If InStr(1, ListWord, CommonLetter1, vbBinaryCompare) = 0 Then Exit Do
               ChkStr = Replace(ListWord, CommonLetter1, "", 1, 1, vbBinaryCompare)
               If InStr(1, ChkStr, CommonLetter2, vbBinaryCompare) = 0 Then Exit Do
            End If
            'Eliminate words with non-matching letters.
            For k = 1 To Len(ListWord)
                    If InStr(1, SampleWord, Mid(ListWord, k, 1), vbBinaryCompare) = 0 Then Exit Do
            'Use each letter only once(if selected).
                    If chkOnce.Value Then
                        SampleWord = Replace(SampleWord, Mid(ListWord, k, 1), "", 1, 1, vbBinaryCompare)
                    End If
            Next k
            'Eliminate target word from list in anagram mode(optional)
            If chkAnagram.Value And EnteredWord = ListWord Then Exit Do
            'the list word qualifies, so we add it to the list box
            lstResults.AddItem ListWord
'*********************************************************************************************
'*********************************************************************************************
            Counter = Counter + 1   'word count
        Exit Do
    Loop
    SampleWord = EnteredWord  'restore target word
Loop
Close #1

If chkAnagram.Value Then
    ModeStr = "Anagrams"
Else
    ModeStr = "words"
End If

lblResults.Caption = Counter & Space(1) & ModeStr & Space(1) & "found in" & Space(1) _
& Format(Timer - StartTime, "0.00") & Space(1) & "Sec."

Me.MousePointer = vbNormal
Exit Sub

ErrHandler:
Select Case Err.Number
    Case 53
        MsgBox "File not found - Error " & Err.Number, vbCritical, "Target Words"
    Case 55
        MsgBox "File already open - Error " & Err.Number, vbCritical, "Target Words"
        Close #1
    Case Else
        MsgBox "Error " & Err.Number & " occured", vbCritical, "Target Words"
End Select
Close #1
End Sub

Private Sub Form_Activate()
txtEnter.SetFocus
End Sub

Private Sub lstResults_Click()

If Left(lstResults.Text, 1) = "<" Then
    lstResults.ToolTipText = ""
    Exit Sub
Else
    lstResults.ToolTipText = lstResults.Text & Space(1) & "(" & Len(lstResults.Text) & " letters)"
End If

End Sub

Private Sub txtCommon1_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 8, 13
        Exit Sub
    Case Is < 65
        KeyAscii = 0
        Exit Sub
    Case 91 To 96
        KeyAscii = 0
        Exit Sub
    Case Is > 122
        KeyAscii = 0
        Exit Sub
End Select
     txtCommon2.SetFocus
    If txtEnter = "" Then
        MsgBox "Enter Target Word", vbExclamation, "Target Words"
        KeyAscii = 0
        txtEnter.SetFocus
        Exit Sub
    End If
    If InStr(1, txtEnter.Text, Chr(KeyAscii), 1) = 0 Then
        MsgBox "Letter not found in Target Word", vbInformation, "Target Words"
        KeyAscii = 0
        txtCommon1.SetFocus
    End If
   
End Sub
Private Sub txtCommon2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 13
        Exit Sub
    Case Is < 65
        KeyAscii = 0
        Exit Sub
    Case 91 To 96
        KeyAscii = 0
        Exit Sub
    Case Is > 122
        KeyAscii = 0
        Exit Sub

End Select
    If InStr(1, txtEnter.Text, Chr(KeyAscii), 1) = 0 Then
        MsgBox "Letter not found in Target Word", vbInformation, "Target Words"
        KeyAscii = 0
        txtCommon2.SetFocus
    End If
End Sub
Private Sub txtEnter_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 8, 13
        Exit Sub
    Case Is < 65
        KeyAscii = 0
        Exit Sub
    Case 91 To 96
        KeyAscii = 0
        Exit Sub
    Case Is > 122
        KeyAscii = 0
        Exit Sub
End Select
End Sub

Private Sub txtEnter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If txtEnter.Text = "" Then
    txtEnter.ToolTipText = "Text Box Empty"
Else
    txtEnter.ToolTipText = txtEnter.Text & " (" & Len(txtEnter.Text) & " letters)"
End If

End Sub

Private Sub txtMinimum_KeyPress(KeyAscii As Integer)

If txtEnter.Text = "" Then
    MsgBox "Target Word must be entered First", vbExclamation, "Target Words"
    KeyAscii = 0
    txtEnter.SetFocus
    Exit Sub
End If

Select Case KeyAscii
    Case 8                  'allow backspace
        Exit Sub
    Case 13, 32
        KeyAscii = 0
        Exit Sub
    Case 48 To 57            'allow only numbers 1 - 9

    Case Is > 57
        KeyAscii = 0
        Exit Sub
       

End Select
   
    If chkOnce.Value And Len(txtEnter.Text) < Chr(KeyAscii) Then
        MsgBox "When 'Use each letter only once' is selected" & vbLf & vbCr & _
        "words cannot have more letters than Target Word", vbInformation, "Target Words"
        KeyAscii = 0
        txtMinimum.SetFocus
    End If
End Sub
Private Sub ClearAll()
With txtEnter
    .Text = ""
    .Enabled = True
    .SetFocus
End With
With txtMinimum
    .Enabled = False
    .Text = ""
End With
With txtCommon1
    .Enabled = False
    .Text = ""
End With
With txtCommon2
    .Enabled = False
    .Text = ""
End With
With chkOnce
    .Enabled = True
    .Value = 0
End With
With chkCommon
    .Enabled = True
    .Value = 0
End With
With chkMinimum
    .Enabled = True
    .Value = 0
End With
With chkAnagram
    .Enabled = True
    .Value = 0
End With
cmdFind.Caption = "Find Words"
lblResults.Caption = ""
lstResults.Clear
End Sub
'*****************************************************************
'*****************************************************************
'For VB5 users. Add this code to a BAS module
'to emulate the VB6 'Replace' function.
'Source: http://support.microsoft.com/kb/188007

'Public Function Replace(sIn As String, sFind As String, _
'            sReplace As String, Optional nStart As Long = 1, _
'            Optional nCount As Long = -1, Optional bCompare As _
'            VbCompareMethod = vbBinaryCompare) As String
'
'          Dim nC As Long, nPos As Integer, sOut As String
'
'          sOut = sIn
'          nPos = InStr(nStart, sOut, sFind, bCompare)
'          If nPos = 0 Then GoTo EndFn:
'          Do
'             nC = nC + 1
'              sOut = Left(sOut, nPos - 1) & sReplace & _
'              Mid(sOut, nPos + Len(sFind))
'              If nCount <> -1 And nC >= nCount Then Exit Do
'              nPos = InStr(nStart, sOut, sFind, bCompare)
'          Loop While nPos > 0
'EndFn:
'          Replace = sOut
'End Function



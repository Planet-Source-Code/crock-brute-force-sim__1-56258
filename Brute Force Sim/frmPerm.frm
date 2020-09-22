VERSION 5.00
Begin VB.Form frmPerm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Brute Force Sim"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8910
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   285
      Left            =   7440
      TabIndex        =   22
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   285
      Left            =   6120
      TabIndex        =   21
      Top             =   3360
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5280
      Top             =   3240
   End
   Begin VB.Frame fraStats 
      Caption         =   "Stats"
      ForeColor       =   &H80000002&
      Height          =   3135
      Left            =   5760
      TabIndex        =   11
      Top             =   120
      Width           =   3015
      Begin VB.Label lblStats 
         Caption         =   "Total Perms:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
         ToolTipText     =   "Total permutations for all loops."
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblPermsTotal 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblPermPerSec 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   26
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblStats 
         Caption         =   "Speed:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "Permutations per second."
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ready"
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label lblStats 
         Caption         =   "Status:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   25
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lblPermsMax 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblPermCur 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblTime 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00:00:00"
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lblLength 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblPassword 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?"
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblStats 
         Caption         =   "Time:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "Elapsed time."
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblStats 
         Caption         =   "Length:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Current length."
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblStats 
         Caption         =   "Password:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Current permutated password."
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblStats 
         Caption         =   "Perms:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Permutation count for current loop."
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblStats 
         Caption         =   "Max Perms:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Maximum permutations for this loop."
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame fraPassword 
      Caption         =   "Password"
      ForeColor       =   &H80000002&
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5535
      Begin VB.TextBox txtMax 
         Height          =   285
         Left            =   3840
         TabIndex        =   31
         Text            =   "8"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtMin 
         Height          =   285
         Left            =   2640
         TabIndex        =   30
         Text            =   "1"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Text            =   "QWERTY"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label 
         Caption         =   "Max Length:"
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label 
         Caption         =   "Min Length:"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label 
         Caption         =   "Enter Password:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame fraCharSet 
      Caption         =   "Character Set"
      ForeColor       =   &H80000002&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   5535
      Begin VB.CheckBox Check 
         Caption         =   "All Keyboard Characters (32 - 126)"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   5
         Top             =   360
         Width           =   2775
      End
      Begin VB.CheckBox Check 
         Caption         =   "Lower Case"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CheckBox Check 
         Caption         =   "Upper Case"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   2535
      End
      Begin VB.CheckBox Check 
         Caption         =   "Numeric"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtCharSet 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   1560
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmPerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Bruteforce As clsPermArr
Dim dSpeed As Double
Dim lTimeElapsed As Long

Private Sub cmdGo_Click()

    Dim fFound As Boolean
    Dim iLengthMin As Integer
    Dim iLengthMax As Integer
    Dim lLengthCur As Long
    
    Dim iBase As Integer
    
    ' Continue?
    If Not ValidInput Then Exit Sub
        
    iLengthMin = Val(txtMin)
    iLengthMax = Val(txtMax)
    iBase = Len(txtCharSet)
            
    ' Give m_Bruteforce some info.
    With m_Bruteforce
        .InitByteSet = txtCharSet
        .InitSoughtArr = txtPassword
        .fActive = True
        lblStatus.Caption = "Busy"
        
    End With
    
    Call ToggleCtls
    
    ' First of many loops!
    For lLengthCur = iLengthMin To iLengthMax
        ' Assign current length to label.
        lblLength.Caption = lLengthCur
        
        ' Clac Max perms for this length.
        lblPermsMax.Caption = iBase ^ lLengthCur
        
        ' Allocate the first password.
        m_Bruteforce.Password = String$(lLengthCur, Left$(txtCharSet, 1))
        
        ' Which permutation function to call?
        If lLengthCur = 1 Then
            fFound = m_Bruteforce.PermByte
            
        Else
            fFound = m_Bruteforce.PermByteArr
            
        End If
        
        If fFound Or Not m_Bruteforce.fActive Then Exit For
        
    Next
        
    ' Permutation process has ended.  Get current info.
    With m_Bruteforce
        .fActive = False    'reset flag, only needed for clean App exit.
        lblPassword.Caption = .Password     ' Get current password.
        lblPermCur.Caption = .PermCount     ' Get current count.
    
    End With
    
    ' Was the password found?
    If fFound Then
        lblStatus.Caption = "Found"
        
    Else
        lblStatus.Caption = "Not Found"
        
    End If
        
    Call ToggleCtls
    
End Sub

Private Sub cmdStop_Click()

    ' Cancel permutation.
    m_Bruteforce.fActive = False
                
End Sub

Private Sub Form_Load()
    
    ' One second interval for timer.
    Timer1.Interval = 1000
    Timer1.Enabled = False
    
    ' Create new instance.
    Set m_Bruteforce = New clsPermArr
    
    ' Create a not to difficult example.  Takes 16 seconds on my P3 800 Mhz.
    txtPassword = "QWERTY"
    txtMin = 1
    txtMax = 8
    Check(1).Value = 1
    Call Check_Click(1)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
         
    ' Unload?
    If m_Bruteforce.fActive Then    'don't unload
        MsgBox "Permutation in progress." & vbNewLine & _
        "Click Stop before exiting.", vbExclamation

        Cancel = True   ' cancel unload.

    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Release from memory.
    Set m_Bruteforce = Nothing
    Set frmPerm = Nothing
    
End Sub

Private Sub Check_Click(Index As Integer)

    ' Create the character set.
    
    Dim sCharSet As String
    Dim i As Long
           
    ' Assign character set to corresponding tag.
    Check(0).Tag = "0123456789"
    Check(1).Tag = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Check(2).Tag = "abcdefghijklmnopqrstuvwxyz"
        
    ' Easier to loop for this charset.
    For i = 32 To 126
        sCharSet = sCharSet & Chr(i)
        
    Next
    
    Check(3).Tag = sCharSet
    
    sCharSet = ""
    
    ' Create CharSet according to the selected Check boxes.
    For i = 0 To Check.Count - 1
        If Check(i).Value = vbChecked Then
            Select Case i
                Case 0 To 2
                    sCharSet = sCharSet & Check(i).Tag
                    
                Case 3
                    sCharSet = Check(i).Tag
                    
            End Select
            
        End If
        
        Check(i).Tag = ""    ' release from memory
        
    Next
   
   ' Assign charset to textbox.
   txtCharSet = sCharSet
      
End Sub

Private Sub txtCharSet_Change()

    Call PermsTotal
    
End Sub

Private Sub txtMax_Change()

    Call PermsTotal
    
End Sub

Private Sub txtMin_Change()

    Call PermsTotal
    
End Sub

Private Function ValidInput() As Boolean

    ' Validates user input.  Attempts to trap any errors before calling PermArr function.
    Dim iLengthMin As Integer
    Dim iLengthMax As Integer
    
    ' First reset some variables.
    dSpeed = 0          ' Reset speed counter
    lTimeElapsed = 0    ' Reset time counter
    
    ' Restore labels default values.
    lblPermsMax.Caption = 0
    lblPermCur.Caption = 0
    lblPermPerSec.Caption = 0
    lblPassword.Caption = "?"
    lblLength.Caption = 0
    lblTime.Caption = "00:00:00"
    lblStatus.Caption = "Ready"
        
    iLengthMin = Val(txtMin)
    iLengthMax = Val(txtMax)

    'Ensure length is valid.
    If iLengthMin < 1 Or iLengthMin > iLengthMax Then
        MsgBox "Cannot continue invalid length."
        GoTo Invalid

    End If
        
    ' Ensure there is a character set to work with.
    If Len(txtCharSet) = 0 Then
        MsgBox "Cannot continue no character set."
        GoTo Invalid
        
    End If
   
   ' Ensure there is a password to find.
    If Len(txtPassword) = 0 Then
        MsgBox "Cannot continue no password to find."
        GoTo Invalid
      
    End If
         
    ' All OK
    ValidInput = True
  
Invalid:

End Function

Private Sub ToggleCtls()

    ' Enable\Disable controls.
    
    fraCharSet.Enabled = Not fraCharSet.Enabled
    fraPassword.Enabled = Not fraPassword.Enabled
    
    cmdGo.Enabled = Not cmdGo.Enabled
    Timer1.Enabled = Not Timer1.Enabled
    
End Sub

Private Sub PermsTotal()

    'Calculate total permutations, put result in a label.
   
    Dim iLengthMin As Integer
    Dim iLengthMax As Integer
    Dim lLengthCur As Long
    Dim iBase As Integer
    Dim dPermsTotal As Double
    
    On Error Resume Next
    
    iLengthMin = Val(txtMin)
    iLengthMax = Val(txtMax)
    iBase = Len(txtCharSet)
    
    ' Compute total permutation with repetitions.
    For lLengthCur = iLengthMin To iLengthMax
        dPermsTotal = dPermsTotal + (iBase ^ lLengthCur)
        
    Next
        
    lblPermsTotal.Caption = dPermsTotal
    
End Sub

Private Sub Timer1_Timer()
    
    Dim sTimeElapsed As String
        
    lTimeElapsed = lTimeElapsed + 1
    
    '24hr stopwatch will reset at 23:59:59.  So calc days if needed.
    sTimeElapsed = "00:00:00"
    
    ' calc hours keeping value within range 0 - 23
    Mid$(sTimeElapsed, 1) = Format$(lTimeElapsed \ 3600 Mod 24, "00")
    
    ' calc minutes keeping value within range 0 - 59
    Mid$(sTimeElapsed, 4) = Format$(lTimeElapsed \ 60 Mod 60, "00")
    
    ' calc seconds keeping value within range 0 - 59
    Mid$(sTimeElapsed, 7) = Format$(lTimeElapsed Mod 60, "00")
    
    ' Assign elapsed time to label.
    lblTime.Caption = sTimeElapsed
    
    ' Get count.
    lblPermCur.Caption = m_Bruteforce.PermCount
    
    ' Get current password.
    lblPassword.Caption = m_Bruteforce.Password
    
    ' Calc permutations per second.
    lblPermPerSec.Caption = m_Bruteforce.PermCount - dSpeed
    
    ' Store count ready for next perms per second.
    dSpeed = m_Bruteforce.PermCount
           
End Sub

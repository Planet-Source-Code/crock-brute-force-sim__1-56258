VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPermArr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Permutation with repetitions algorithm.
' Author:crock@gofrontier.co.uk
' Date 18/09/04
' This class permutates a byte array of a given size using a given byte set.
' Note "repetitions" means the same character can be repeated in the permutation
' not that there are repeats of the permutation.
' Compile before testing it will be slow in the IDE.
' Use VB5 to compile if possible, VB6 is slightly slower.

Option Explicit

'******* For faster DoEvents
Private Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long
Private Const QS_MOUSEBUTTON As Long = &H4
Private Const QS_PAINT As Long = &H20
Private Const QS_TIMER As Long = &H10
'******* For faster DoEvents

Private Declare Sub CopyMemory Lib "kernel32" Alias _
    "RtlMoveMemory" (dest As Any, source As Any, _
    ByVal numBytes As Long)

Private m_bSoughtArr() As Byte     ' known password
Private m_iSoughtLen As Integer    ' known length

Private m_bByteSet() As Byte
Private m_bCountArr() As Byte
Private m_bPwrdArr() As Byte
Private m_iPwrdLen As Integer
Private m_iBase As Integer

Public fActive As Boolean

Public Function PermByte() As Boolean

    ' Had to write this just to cope with a password length of one!  I didn't want to
    ' include error trapping/If statement in the main loop.
    
    Dim lPos As Long
    Dim iFound As Integer
    Dim iCurLen As Integer
    
    ' Align length to ubound.
    iCurLen = m_iPwrdLen - 1
    
    For lPos = 0 To m_iBase - 1
        ' Change last byte on every pass.  The only byte in this case.
        m_bPwrdArr(iCurLen) = m_bByteSet(lPos)
        
        '****** Test for password match.  ******
        iFound = m_iSoughtLen
        
        If m_bPwrdArr(iCurLen) = m_bSoughtArr(iCurLen) Then iFound = iFound - 1

        If iFound = 0 Then GoTo PasswordFound
        '***************************************
                                                                
    Next
    
PasswordNotFound:
    Exit Function

PasswordFound:
    PermByte = True
    
End Function

Friend Function PermByteArr() As Boolean
    
    ' Warning password length most be > 1 to use this function.
    
    Dim lPos As Long
    Dim lTemp As Long   ' Temp var gets reused several times.
    Dim iCurLen As Integer
    Dim iFound As Integer
        
    ' Align length to ubound.
    iCurLen = m_iPwrdLen - 1
      
    'start the loop
    Do
        For lPos = 0 To m_iBase - 1
            ' Change last byte on every pass.
            m_bPwrdArr(iCurLen) = m_bByteSet(lPos)
                        
            '****** Test for password match.  ******
            ' In most brute force examples the test for password match usually
            ' consists of "if StringGenerated = StringPassword then" this uses a
            ' loop to compare each byte.  It is important to process all the bytes
            ' for a true simulation.  We could of cause exit the loop as soon as
            ' the "byteGenerated <> bytePassword" but that would be cheating!
            ' It may appear to some that using the known length (m_iSoughtLen) of
            ' the sought password is cheating.  This is not the case, no advantage
            ' is gained using this value it is used to maintain the integrity of
            ' the simulation.
            ' Bare in mind changes here only effect the simulation speed!  As such
            ' should not be seen as a way of improving the algorithm.
            
            iFound = m_iSoughtLen

            For lTemp = 0 To iCurLen
                If m_bPwrdArr(lTemp) = m_bSoughtArr(lTemp) Then iFound = iFound - 1

            Next

            If iFound = 0 Then GoTo PasswordFound
            '***************************************
                                                                
        Next
        
        ' Base counter loop.  Change other bytes?  Will be at least one to change.
        For lPos = iCurLen - 1 To 0 Step -1
            lTemp = m_bCountArr(lPos) + 1
            
            If lTemp = m_iBase Then  'carry
                lTemp = 0    ' On the last pass this value is used to stop the main loop.
                m_bCountArr(lPos) = lTemp   'reset
                m_bPwrdArr(lPos) = m_bByteSet(lTemp)
                                
                If GetQueueStatus(QS_MOUSEBUTTON Or QS_PAINT Or QS_TIMER) Then
                    DoEvents
                    If Not fActive Then GoTo PasswordNotFound   ' Cancel?
                    
                End If
                               
            Else
                m_bCountArr(lPos) = lTemp    'increment
                m_bPwrdArr(lPos) = m_bByteSet(lTemp)
                Exit For    ' nothing to carry so bail out.

            End If
        
        Next
          
    Loop Until lTemp = 0
       
PasswordNotFound:
    Exit Function

PasswordFound:
    PermByteArr = True
    
End Function

Public Property Let InitByteSet(ByVal NewVal As String)

    ' initialise m_bByteSet with charset
    
    m_iBase = Len(NewVal)

    ReDim m_bByteSet(0 To m_iBase - 1)

    CopyMemory m_bByteSet(0), ByVal NewVal, m_iBase
    
End Property

Public Property Let InitSoughtArr(ByVal NewVal As String)

    ' Initialise m_bSoughtArr with the password to find.
    
    m_iSoughtLen = Len(NewVal)

    ReDim m_bSoughtArr(0 To m_iSoughtLen - 1)
    
    CopyMemory m_bSoughtArr(0), ByVal NewVal, m_iSoughtLen
    
End Property

Public Property Let PermCount(ByVal NewVal As Double)
    
    ' Could be used to load a previous permutation.
    ' would need to initialise m_bCountArr with NewVal
    ' Would need the current length, should be able to calculate
    ' it from NewVal as m_iPwrdLen = ubound(m_bCountArr) + 2.
    
    ' Would need m_bByteSet that was used.
    ' iBase = ubound(m_bByteSet) + 1
    
End Property

Public Property Get PermCount() As Double

    ' Compute count from m_bCountArr by converting the base count back to base 10.
    ' Accurate to within the givin base value as the count array does not
    ' store the units.
 
    Dim i As Long
    Dim iPow As Integer
    
    On Error GoTo NoCount    ' Necessary for password length of one.
    
    iPow = UBound(m_bCountArr) + 1

    For i = 0 To UBound(m_bCountArr)
        PermCount = PermCount + m_bCountArr(i) * (m_iBase ^ iPow)
        iPow = iPow - 1

    Next
         
    Exit Property
    
NoCount:
    PermCount = 0
    
End Property

Public Property Let Password(ByVal NewVal As String)

    ' Initialise password array with the first password.
    On Error Resume Next    ' Necessary for password length of one.
    
    m_iPwrdLen = Len(NewVal)

    ReDim m_bPwrdArr(0 To m_iPwrdLen - 1)
    
    CopyMemory m_bPwrdArr(0), ByVal NewVal, m_iPwrdLen
    
    'Dimension count array while we are here.  Always one less than m_bPwrdArr.
    ReDim m_bCountArr(0 To m_iPwrdLen - 2)
    
End Property

Public Property Get Password() As String
    
    ' Convert m_bPwrdArr to a string.
    
    Password = Space$(m_iPwrdLen)
    
    CopyMemory ByVal Password, m_bPwrdArr(0), m_iPwrdLen
        
End Property
 
Private Sub Class_Terminate()
    
    'not sure if this has got to be done but it does no harm.
    Erase m_bByteSet, m_bCountArr, m_bPwrdArr, m_bSoughtArr

End Sub

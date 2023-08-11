VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDriveMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class: CDriveMapper
'Author: MildewManOne from Reddit
'Purpose: Maps and Unmaps drives on the user's computer.


'If not using late binding, project requires the
'"Windows Script Host Object Model" reference for WshNetwork Object
#Const bLateBind = 1

Private Enum MappingTypes
    MAP_DRIVE
    UNMAP_DRIVE
End Enum

Private m_mappedDrives As Collection

'If this is True (default), any newly mapped drives will be unmapped when class is destroyed.
'Use UnmapOnTerminate Property to change this value.
Private m_unmapDrives As Boolean

#If bLateBind = 1 Then
    Private m_fso As Object
#Else
    Private m_fso As Scripting.FileSystemObject
#End If
Option Explicit
Private Sub Class_Initialize()
    m_unmapDrives = True
    Set m_mappedDrives = New Collection
    Call InitFileSysObject
    Debug.Assert (Not m_mappedDrives Is Nothing)
End Sub
Private Sub Class_Terminate()
    On Error Resume Next    'Make sure no errors thrown during class destruction.
    
    If m_unmapDrives Then
        Call UnmapAllDrives
    End If
    Set m_mappedDrives = Nothing
    
    Set m_fso = Nothing
    On Error GoTo 0
End Sub
Private Function InitFileSysObject() As Boolean
    If m_fso Is Nothing Then
#If bLateBind = 1 Then
        Set m_fso = CreateObject("Scripting.FileSystemObject")
#Else
        Set m_fso = New Scripting.FileSystemObject
#End If
    End If
    
    InitFileSysObject = Not m_fso Is Nothing
End Function
Public Property Let UnmapOnTerminate(ByVal bUnmap As Boolean)
    m_unmapDrives = bUnmap
End Property
Public Property Get UnmapOnTerminate() As Boolean
    UnmapOnTerminate = m_unmapDrives
End Property
Private Function UnmapAllDrives()
    Dim count As Long
    count = m_mappedDrives.count
    
    If count <= 0 Then
        Exit Function
    End If
    
    On Error Resume Next
    Dim i As Long

    For i = count To 1 Step -1
        Call UnMapDrive(CStr(m_mappedDrives.Item(i)))
    Next i
    
    On Error GoTo 0
    
End Function
Private Function FormatDriveLetter(ByVal driveLetter As String, Optional ByVal addColon As Boolean = True) As String

    Dim length As Long
    Const comparison As String = "[A-Z]"
    Dim retval As String
    Dim tmp As String
    
    If driveLetter = vbNullString Then
        Exit Function
    Else
        length = Len(driveLetter)
        If length = 1 Then
	    'make sure the drive letter is uppercase
            retval = UCase(driveLetter)

            'If driveLetter is only 1 character long, but is not a letter a-z, then it's not a valid drive letter.
            If Not retval Like comparison Then
                Exit Function
            End If

            'Add a colon to the end.
            If addColon Then
                retval = retval & ":"
            End If
        Else
            tmp = UCase(Left(driveLetter, 1))
            'if length > 1, check to see if maybe the user passed a path (ex: "G:\path") instead of a letter.
            'or if they just passed some random string that happens to begin with a drive letter (ex: "Go to sleep")
            If tmp Like comparison And Mid(driveLetter, 2, 1) = ":" Then
                retval = tmp
                If addColon Then
                    retval = retval & ":"
                End If
            End If
        End If
    End If
    
    FormatDriveLetter = retval
    
End Function
Private Function DoDriveMapping(ByVal driveLetter As String, ByVal driveAddress As String, ByVal command As MappingTypes, Optional ByVal forceDuplicateMapping As Boolean = False) As Long
    
    On Error GoTo ErrHandler
    
    Dim hr As Long
    Dim localDriveLetter As String

#If bLateBind = 1 Then
    Dim WNet As Object
#Else
    Dim WNet As IWshRuntimeLibrary.WshNetwork   'Requires "Windows Script Host Object Model" reference
#End If
    
    hr = 0
    
    If driveLetter = vbNullString Or (command = MAP_DRIVE And driveAddress = vbNullString) Then
        hr = -1
    ElseIf command = UNMAP_DRIVE Then
	If (IsDriveAvailable(driveLetter)) Then
	    hr = 1	'If drive is available, it's already unmapped.
	End If
    Else
        'the drive letter has to have a colon after it when passing to WshNetwork (i.e. "D:")
        localDriveLetter = FormatDriveLetter(driveLetter)
        If localDriveLetter = vbNullString Then
            'An invalid string was passed to driveLetter.
            hr = -1
        End If
    End If
    
    If hr <> 0 Then
        GoTo CleanUp
    End If

#If bLateBind = 1 Then
    Set WNet = CreateObject("WScript.Network")
#Else
    Set WNet = New IWshRuntimeLibrary.WshNetwork
#End If

    Debug.Assert (Not WNet Is Nothing)
    
    Select Case command
        Case MAP_DRIVE
            'Check to make sure the drive is not already mapped to something else.
            If (IsDriveAvailable(driveLetter)) Then
                'If the user didn't specify that duplicate mappings were ok,
                'make sure the address is not already mapped to anothier drive letter.
                If Not forceDuplicateMapping And (FindDriveLetterByAddress(driveAddress, vbLong) <> 0) Then
                    'Address already mapped. Don't continue.
                    Debug.Print "CDriveMapper::DoDriveMapping - driveAddress already mapped to " & FindDriveLetterByAddress(driveAddress, vbString)
                    hr = 1
                    GoTo CleanUp
                End If
                Call WNet.MapNetworkDrive(localDriveLetter, driveAddress)
		m_mappedDrives.Add driveLetter	'Add the newly mapped drive letter to the collection.
            Else
                Debug.Print "CDriveMapper::DoDriveMapping - driveLetter already mapped to something else."
                hr = 1
                GoTo CleanUp
            End If
        Case UNMAP_DRIVE
            Call WNet.RemoveNetworkDrive(localDriveLetter)
    End Select
    
    hr = 0
    
CleanUp:
    On Error Resume Next
    Set WNet = Nothing
    DoDriveMapping = hr
    On Error GoTo 0
    Exit Function
ErrHandler:
    Debug.Print "CDriveMapper::DoDriveMapping Error Thrown"
    Debug.Print "Error number = " & CStr(Err.Number)
    Debug.Print "Error Description = " & Err.Description
    hr = -1
    Resume CleanUp
    
End Function
Public Function MapDrive(ByVal driveAddress As String, Optional ByVal driveLetterAscii As Long = 0, Optional ByVal forceIfAddressMapped As Boolean = False) As Long
    If Not IsValidDriveLetter(driveLetterAscii) Then
        Call FindAvailableDriveAscii(driveLetterAscii)
        'Call FindAvailableDriveLetter(driveLetter)
        'If driveLetter = vbNullString Then
        If Not IsValidDriveLetter(driveLetterAscii) Then
            MapDrive = -1
            Exit Function
        End If
    End If
    MapDrive = DoDriveMapping(Chr(driveLetterAscii), driveAddress, MAP_DRIVE, forceIfAddressMapped)
End Function
Private Function IsValidDriveLetter(ByVal ascii As Long) As Boolean
    IsValidDriveLetter = (ascii >= Asc("A") And ascii <= Asc("Z"))
End Function
Public Function UnMapDrive(Optional ByVal driveLetterAscii As Long = 0, Optional ByVal driveAddress As String = vbNullString) As Long
    If IsValidDriveLetter(driveLetterAscii) Then
        UnMapDrive = DoDriveMapping(Chr(driveLetterAscii), vbNullString, UNMAP_DRIVE)
    ElseIf driveAddress <> vbNullString Then
        driveLetterAscii = FindDriveLetterByAddress(driveAddress, vbLong)
        If driveLetterAscii <> 0 Then
            UnMapDrive = Me.UnMapDrive(Chr(driveLetterAscii))
        End If
    Else
        UnMapDrive = -1
    End If
End Function

Private Function RemoveTrailingSlash(ByVal driveAddress As String) As String
    
    Dim lastChar As String
    
    lastChar = Right(driveAddress, 1)
    
    If lastChar = "\" Or lastChar = "/" Then
        driveAddress = Left(driveAddress, Len(driveAddress) - 1)
    End If
    
    RemoveTrailingSlash = driveAddress

End Function
Private Function FindDriveLetterByAddress(ByVal driveAddress As String, Optional ByVal returnType As VbVarType) As Variant
'Attempts to find a drive letter that driveAddress is mapped to on the user's computer.
'If successful, the function will return the Drive Letter where the address is currently mapped.
'If unsuccessful, the function returns "0". (zero)
    
    If returnType = vbString Then
        FindDriveLetterByAddress = "0"
    Else
        FindDriveLetterByAddress = CLng(0)
    End If

    If driveAddress = vbNullString Or Not InitFileSysObject() Then
        Exit Function
    End If
    
    Dim d As Object
    
    driveAddress = RemoveTrailingSlash(driveAddress)

    For Each d In m_fso.drives
        Debug.Print d.driveLetter & ": = " & d.ShareName
        'Check to see if the drive's ShareName Property is not vbNullString
        If d.ShareName <> vbNullString Then
            'Now check to see if the drive's ShareName Property matches driveAddress
            If StrComp(driveAddress, d.ShareName, vbTextCompare) = 0 Then
                'Found a match - return this drive's letter.
                If returnType = vbString Then
                    FindDriveLetterByAddress = d.driveLetter
                Else
                    FindDriveLetterByAddress = CLng(Asc(d.driveLetter))
                End If
                Exit For
            End If
        End If
    Next d

End Function
Private Function IsDriveAvailable(ByVal driveLetter As String) As Boolean
    If driveLetter = vbNullString Or Not InitFileSysObject() Then
        Exit Function
    End If

    Dim d As Object

    IsDriveAvailable = True     'Assume it's available and the loop below will prove if it's not.
    
    For Each d In m_fso.drives
        If driveLetter = d.driveLetter Then
            IsDriveAvailable = False
            Exit For
        End If
    Next
    
    Set d = Nothing

End Function
Private Function FindAvailableDriveAscii(ByRef out_ascii As Long) As Long
    Dim driveLetter As String
    Dim tmp As Long
    
    out_ascii = 0
    
    FindAvailableDriveAscii = FindAvailableDriveLetter(driveLetter)
    If FindAvailableDriveAscii = 0 Then
        tmp = Asc(UCase(driveLetter))
        If IsValidDriveLetter(tmp) Then
            out_ascii = tmp
        End If
    End If
End Function
Private Function FindAvailableDriveLetter(ByRef out_DriveLetter As String) As Long
'This function starts at drive letter G and checks each letter through Z to find a letter that is not mounted yet.
'When it finds an open drive letter, it will set the provided DriveLetter variable to that letter and return S_OK.
'If it cannot find an open drive, then it will return S_FALSE and DriveLetter will be vbNullString.

    Dim driveNumber As Long
    Dim hr As Long
    
    If Not InitFileSysObject() Then
        hr = -1
        GoTo CleanUp
    End If

    driveNumber = Asc("Z")		'Operating system tends to assign new drive letters going forward, so we'll start at drive Z and go backwards. 
    out_DriveLetter = vbNullString
    
    hr = 1
    
    Do While (driveNumber >= Asc("G"))			'drives A-F are frequently used. If can't find an open drive b/w G-Z, something is wrong.
        If IsDriveAvailable(Chr(driveNumber)) Then
            out_DriveLetter = Chr(driveNumber)
            hr = 0
            GoTo CleanUp
        End If
        driveNumber = driveNumber - 1
    Loop
    
CleanUp:
    On Error Resume Next
    FindAvailableDriveLetter = hr
    On Error GoTo 0
    Exit Function
ErrHandler:
    hr = -1
    Resume CleanUp
    
End Function

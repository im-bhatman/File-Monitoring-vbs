On Error Resume Next
Set objWbemlocator = CreateObject("WbemScripting.SWbemLocator")

If Err.Number Then
    Wscript.Echo vbCrLf & "Error # " & _
            " " & Err.Description
End If
On Error GoTo 0

Select Case Wscript.Arguments.Count
    Case 1
        'for local machine
        strComputer = Wscript.Arguments.Item(0)
        Set wbemServices = objWbemlocator.ConnectServer _
                        (strComputer,"Root\CIMV2")
    
    Case 3
        'for remote machine
        strComputer = Wscript.Arguments.Item(0)
        strUsername = Wscript.Arguments.Item(1)
        strPassword = Wscript.Arguments.Item(2)
        Set wbemServices = objWbemlocator.ConnectServer _
                        (strComputer,"ROOT\CIMV2",strUsername,strPassword)
    
    Case Else
        Wscript.Quit(0)
End Select


'take file full path from the used
Do
    strFile = InputBox _
        ("please enter a full path for the file to monitor:","Monitor File")
    If strFile = "" Then
        Wscript.Echo "You must enter a file name"
    Else
        Wscript.Echo strFile
    End If
loop While strFile = ""

sizeLimit = 2000 'this will be the size limit for the file.

'Get file Size
Set colFiles = wbemServices.ExecQuery _
        ("Select * from CIM_Datafile where Name='" & strFile &"'")

For Each objFile in colFiles
    strOriginalSize  = objFile.FileSize
    Wscript.Echo "Original Size of the File :: " & strOriginalSize
Next

Wscript.Echo ":::  Monitoring For the File Starting  :::"

Do
    Wscript.Sleep 6000 'Sleep time can be modified as per the requirement
    Set colFiles = wbemServices.ExecQuery _
        ("Select * from CIM_Datafile where Name='" & strFile & "'")
    
    'Get the latest file size
    For Each objFile in colFiles
        strLatestSize = objFile.FileSize
        Wscript.Echo "Latest Size of File ::" & strLatestSize
    Next
    Wscript.Echo "strLatestSize :: " & strLatestSize
    Wscript.Echo "sizeLimit :: " & sizeLimit

    ' if the latestSize of file is  > sizeLimit, send and alret messag to the user
    If CInt(strLatestSize) > sizeLimit Then
        Wscript.Echo "Alert: File Size More than threshold level :: " & strLatestSize
    Else
        Wscript.Echo "File Size is Under Limit"
    End If
Loop
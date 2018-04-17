' -----------------------------------------------------------------------------
' VB Scripts.
' unittest App functions
'
' @category  VB Scripts
' @package   Includes
' @version   20170905
' @copyright (©) 2017, Olivier Jullien <https://github.com/ojullien>
' -----------------------------------------------------------------------------

Option Explicit

' -----------------------------------------------------
' Usefull function
' -----------------------------------------------------

Function toLiteral(x)

  Select Case VarType(x)
    Case vbEmpty
      toLiteral = "<Empty>"
    Case vbNull
      toLiteral = "<Null>"
    Case vbObject
      toLiteral = "<" & TypeName(x) & " object>"
    Case vbString
      toLiteral = qq(x)
    Case Else
      toLiteral = CStr(x)
  End Select

End Function

' -----------------------------------------------------
' string.vbs
' -----------------------------------------------------

Function test_isString ()
    notice "Test string.vbs::isString"
    Dim x
    For Each x In m_APP_UNITTEST_VALUES01
        WScript.Echo toLiteral(x), CStr( isString(x) )
    Next
End Function

Function test_isStringAndNotEmpty ()
    notice "Test string.vbs::isStringAndNotEmpty"
    Dim x
    For Each x In m_APP_UNITTEST_VALUES01
        WScript.Echo toLiteral(x), CStr( isStringAndNotEmpty(x) )
    Next
End Function

' -----------------------------------------------------
' filesystem.vbs
' -----------------------------------------------------

Function test_pwd ()
    notice "Test filesystem.vbs::pwd"
    notice "pwd is : " & pwd
End Function

Function getFileNameTest ( x )
    Dim sReturn : sReturn = "On error"
    On Error Resume Next
    Err.Clear
    sReturn = getFileName(x)
    If Err.Number <> 0 Then
        error "NOK code: " & CStr(Err.Number) & " Description: " & Err.Description
        getFileNameTest = "On error"
    Else
        getFileNameTest = sReturn
    End If
    Err.Clear
    On Error GoTo 0
End Function

Function test_getFileName ()
    notice "Test filesystem.vbs::getFileName"
    Dim x, sReturn
    For Each x In m_APP_UNITTEST_VALUES02
        notice "filename of " & toLiteral(x) & " is : "
        sReturn = getFileNameTest(x)
        notice sReturn
    Next
End Function

Function existsFolderTest ( x )
    Dim bReturn : bReturn = False
    On Error Resume Next
    Err.Clear
    bReturn = existsFolder(x)
    If Err.Number <> 0 Then
        error "NOK code: " & CStr(Err.Number) & " Description: " & Err.Description
        existsFolderTest = False
    Else
        existsFolderTest = bReturn
    End If
    Err.Clear
    On Error GoTo 0
End Function

Function test_existsFolder ()
    notice "Test filesystem.vbs::existsFolder"
    Dim x, bReturn, sPath
    For Each x In m_APP_UNITTEST_VALUES01
        notice "Does folder " & toLiteral(x) & " exist : "
        bReturn = existsFolderTest(x)
        notice CStr(bReturn)
    Next
    sPath = pwd
    notice "Does folder " & toLiteral(sPath) & " exist : " & CStr(existsFolderTest(sPath))
End Function

Function existsFileTest ( x )
    Dim bReturn : bReturn = False
    On Error Resume Next
    Err.Clear
    bReturn = existsFile(x)
    If Err.Number <> 0 Then
        error "NOK code: " & CStr(Err.Number) & " Description: " & Err.Description
        existsFileTest = False
    Else
        existsFileTest = bReturn
    End If
    Err.Clear
    On Error GoTo 0
End Function

Function test_existsFile ()
    notice "Test filesystem.vbs::existsFile"
    Dim x, bReturn, sPath
    For Each x In m_APP_UNITTEST_VALUES02
        notice "Does file " & toLiteral(x) & " exist : "
        bReturn = existsFileTest(x)
        notice CStr(bReturn)
    Next
End Function

Function cleanDirectoryTest ( x )
    Dim bReturn : bReturn = False
    On Error Resume Next
    Err.Clear
    bReturn = cleanDirectory(x)
    If Err.Number <> 0 Then
        error "NOK code: " & CStr(Err.Number) & " Description: " & Err.Description
        cleanDirectoryTest = False
    Else
        cleanDirectoryTest = bReturn
    End If
    Err.Clear
    On Error GoTo 0
End Function

Function test_cleanDirectory ()
    notice "Test filesystem.vbs::cleanDirectory"

    Dim x, bReturn, sPath
    For Each x In m_APP_UNITTEST_VALUES02
        notice "Clean " & toLiteral(x)
        bReturn = cleanDirectoryTest(x)
        notice "cleanDirectory returns: " &CStr(bReturn)
    Next

    On Error Resume Next
    Err.Clear

    Dim oFSO, oFile
    Set oFSO = CreateObject( "Scripting.FileSystemObject" )

    sPath = "c:\Temp\unittest"
    If Not oFSO.FolderExists(sPath) Then oFSO.CreateFolder(sPath)
    notice "Just created " & sPath

    sPath = sPath & "\file.txt"
    Set oFile = oFSO.CreateTextFile( sPath, True, True )
    oFile.Close
    notice "Just created " & sPath

    sPath = "c:\Temp\unittest\a"
    If Not oFSO.FolderExists(sPath) Then oFSO.CreateFolder(sPath)
    notice "Just created " & sPath

    sPath = sPath & "\file-a.txt"
    Set oFile = oFSO.CreateTextFile( sPath, True, True )
    oFile.Close
    notice "Just created " & sPath

    sPath = "c:\Temp\unittest\a\b"
    If Not oFSO.FolderExists(sPath) Then oFSO.CreateFolder(sPath)
    oFSO.CreateFolder(sPath)
    notice "Just created " & sPath

    sPath = sPath & "\file-b.txt"
    Set oFile = oFSO.CreateTextFile( sPath, True, True )
    oFile.Close
    notice "Just created " & sPath

    sPath = "c:\Temp\unittest"
    notice "Clean " & CStr(sPath)
    bReturn = cleanDirectory(sPath)
    If Err.Number <> 0 Then
        error "NOK code: " & CStr(Err.Number) & " Description: " & Err.Description
    End If

    notice "cleanDirectory returns: " &CStr(bReturn)

    Err.Clear
    On Error GoTo 0

End Function

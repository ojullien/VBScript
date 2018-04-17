' -----------------------------------------------------------------------------
' VB Scripts.
' File System functions
'
' @category  VB Scripts
' @package   Includes
' @version   20170911
' @copyright (©) 2017, Olivier Jullien <https://github.com/ojullien>
' -----------------------------------------------------------------------------

Option Explicit

' -----------------------------------------------------------------------------
' Constants
' -----------------------------------------------------------------------------

Const DELETEREADONLY = TRUE

' -----------------------------------------------------------------------------
' Returns the full pathname of the current working directory
' -----------------------------------------------------------------------------

Function pwd ()
    pwd = Replace(WScript.ScriptFullName, "\" & WScript.ScriptName, "")
End Function

' -----------------------------------------------------
' Extract file name
' -----------------------------------------------------

Function getFileName ( ByVal sPath )
    If Not isStringAndNotEmpty(sPath) Then
        error "Usage: getFileName <path as string>"
        Err.raise 5
    End If
    Dim aElements : aElements = Split( trim(sPath), "\" )
    getFileName = aElements( Ubound(aElements) )
End Function

' -----------------------------------------------------------------------------
' Directories
' https://msdn.microsoft.com/de-de/library/1c87day3(v=vs.84).aspx
' -----------------------------------------------------------------------------

' -----------------------------------------------------------------------------
' Checks whether a folder exists
' -----------------------------------------------------------------------------

Function existsFolder ( ByVal sPath )
    If Not isStringAndNotEmpty(sPath) Then
        error "Usage: existsFolder <path as string>"
        Err.raise 5
    End If
    Dim oFSO
    Set oFSO = CreateObject( "Scripting.FileSystemObject" )
    existsFolder = oFSO.FolderExists(sPath)
End Function

Function checkDir ( ByVal sTxt, ByVal sPath )
    If Not isString(sTxt) Then
        error "Usage: checkDir <label as string> <path as string>"
        Err.raise 5
    End If
    noticel sTxt
    If existsFolder(sPath) Then
        success "EXISTS"
        checkDir = True
    Else
        error "MISSING"
        checkDir = False
    End If
End Function

Function cleanDirectory (ByVal sPath)
    If Not isStringAndNotEmpty(sPath) Then
        error "Usage: cleanDirectory <path as string>"
        Err.raise 5
    End If

    On Error Resume Next
    Err.Clear
    cleanDirectory = False

    Dim oFSO
    Set oFSO = CreateObject( "Scripting.FileSystemObject" )
    If oFSO.FolderExists(sPath) Then

        noticel "Cleaning files from " & qq(sPath) & " : "
        oFSO.DeleteFile( sPath & "\*" ), DELETEREADONLY
        If Err.Number <> 0 Then
            error "NOK code: " & CStr(Err.Number) & " Description: " & Err.Description
            Err.Clear
        Else
            success "OK"
            cleanDirectory = True
        End If

        noticel "Cleaning sub folder from " & qq(sPath) & " : "
        oFSO.DeleteFolder( sPath & "\*" ),DELETEREADONLY
        If Err.Number <> 0 Then
            error "NOK code: " & CStr(Err.Number) & " Description: " & Err.Description
            cleanDirectory = False
        Else
            success "OK"
            cleanDirectory = cleanDirectory And True
        End If

    End If

    Err.Clear
    On Error GoTo 0

End Function

Function createDirectory (ByVal sPath)
    If Not isStringAndNotEmpty(sPath) Then
        error "Usage: createDirectory <path as string>"
        Err.raise 5
    End If

    noticel "Creates the folder " & qq(sPath) & " : "
    createDirectory = True

    On Error Resume Next
    Err.Clear

    Dim oFSO
    Set oFSO = CreateObject( "Scripting.FileSystemObject" )
    If oFSO.FolderExists( sPath ) Then
        Success "Folder already exists."
    Else
        fso.CreateFolder( sPath )
        If Err.Number <> 0 Then
            error "NOK code: " & CStr(Err.Number) & " Description: " & Err.Description
            createDirectory = False
        Else
            success "OK"
        End If
    End If

    Err.Clear
    On Error GoTo 0

End Function

Function deleteDirectory (ByVal sPath)
    If Not isStringAndNotEmpty(sPath) Then
        error "Usage: deleteDirectory <path as string>"
        Err.raise 5
    End If

    noticel "Deletes the specified folder " & qq(sPath) & " and its contents : "
    deleteDirectory = True

    On Error Resume Next
    Err.Clear

    Dim oFSO
    Set oFSO = CreateObject( "Scripting.FileSystemObject" )
    If Not oFSO.FolderExists( sPath ) Then
        Success "Folder does not exist."
    Else
        fso.DeleteFolder( sPath )
        If Err.Number <> 0 Then
            error "NOK code: " & CStr(Err.Number) & " Description: " & Err.Description
            deleteDirectory = False
        Else
            success "OK"
        End If
    End If

    Err.Clear
    On Error GoTo 0

End Function

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

Function getFilesDirectory(ByVal sPath, ByRef aFiles)
    If Not isStringAndNotEmpty(sPath) Then
        error "Usage: getFilesDirectory <path as string> <files as ArrayList>"
        Err.raise 5
    End If
    'TODO
    getFilesDirectory = True
End Function

' -----------------------------------------------------------------------------
' Files
'https://msdn.microsoft.com/de-de/library/314cz14s(v=vs.84).aspx
' -----------------------------------------------------------------------------

' -----------------------------------------------------------------------------
' Checks whether a file exists
' -----------------------------------------------------------------------------

Function existsFile ( ByVal sPath )
    If Not isStringAndNotEmpty(sPath) Then
        error "Usage: existsFile <path as string>"
        Err.raise 5
    End If
    Dim oFSO
    Set oFSO = CreateObject( "Scripting.FileSystemObject" )
    existsFile = oFSO.FileExists(sPath)
End Function

Function checkFile ( ByVal sTxt, ByVal sPath )
    If Not isString(sTxt) Then
        error "Usage: checkFile <label as string> <path as string>"
        Err.raise 5
    End If
    noticel sTxt
    If existsFile(sPath) Then
        success "EXISTS"
        checkFile = True
    Else
        error "MISSING"
        checkFile = False
    End If
End Function

Function openFile ( ByVal sPath )
    If Not isStringAndNotEmpty(sPath) Then
        error "Usage: existsFile <path as string>"
        Err.raise 5
    End If
    Dim oFSO
    Set oFSO = CreateObject( "Scripting.FileSystemObject" )
    existsFile = oFSO.FileExists(sPath)
End Function

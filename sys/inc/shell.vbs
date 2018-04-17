' -----------------------------------------------------------------------------
' VB Scripts.
' Shell functions
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

Const WAITONRETURN = True
Const WINDOWSTYLE = 0

' -----------------------------------------------------------------------------
' Runs a program in a new process.
' -----------------------------------------------------------------------------

Function runCommand ( ByVal sCommand )
    Dim oShell, iErrorCode

    If Not isStringAndNotEmpty(sCommand) Then
        error "Usage: runCommand <command as string> "
        Err.raise 5
    End If

    noticel "Run the command via shell:"

    On Error Resume Next
    Err.Clear

    Set oShell = WScript.CreateObject("WScript.Shell")
    iErrorCode = oShell.Run( sCommand, WINDOWSTYLE, WAITONRETURN )

    If Err.Number <> 0 Then
        error "NOK code: " & CStr(Err.Number) & " (0x" & hex(Err.Number) & ") Description: " & Err.Description
    Else
        success "OK code: " & CStr(iErrorCode)
    End If

    runCommand = iErrorCode

    Err.Clear
    On Error GoTo 0

End Function

' -----------------------------------------------------------------------------
' Runs an application in a child command-shell, providing access to the
' StdIn/StdOut/StdErr streams.
' -----------------------------------------------------------------------------

Function contiger ( ByVal sPath )
    Dim sBin : sBin = "C:\Program Files\SysinternalsSuite\Contig.exe"
    Dim sOption : sOption = "-s -q -nobanner"
    Dim sCommand : sCommand = qq(sBin) & " " & sOption & " " & qq(sPath & "\*")
    Dim sStdOut : sStdOut = ""
    Dim sStdErr : sStdErr = ""
    Dim oShell, oExec

    contiger = False
    On Error Resume Next
    Err.Clear

    notice "Contiger:"
    notice vbTab & "Path: " & sPath
    noticel vbTab & "Exit returns: "

    Set oShell = WScript.CreateObject ("WScript.Shell")
    set oExec = oShell.Exec(  sCommand )
    If Err.Number <> 0 Then
        error "NOK code: " & CStr(Err.Number) & " (0x" & hex(Err.Number) & ") Description: " & Err.Description
    Else
        success "OK code: " & CStr( oExec.ExitCode )
        Do While oExec.Status = 0
            Do While Not oExec.StdOut.AtEndOfStream
                sStdOut = sStdOut & oExec.StdOut.ReadLine & vbCrLf
            Loop
            Do While Not oExec.StdErr.AtEndOfStream
                sStdErr = sStdErr & oExec.StdErr.ReadLine & vbCrLf
            Loop
            WScript.Sleep 0
        Loop
        notice sStdOut & sStdErr
        contiger = True
    End If

    Err.Clear
    On Error GoTo 0

End Function

Function cleanDirectory (ByVal sPath)
    Dim oFSO
    Set oFSO = CreateObject( "Scripting.FileSystemObject" )
    On Error Resume Next
    Err.Clear
    noticel "Cleaning " & sPath & ": "
    oFSO.DeleteFile( sPath & "\*" ), DeleteReadOnly
    If Err.Number <> 0 Then
        error "NOK code: " & CStr(Err.Number) & " Description: " & Err.Description
        cleanDirectory = False
    Else
        success "OK"
        cleanDirectory = True
    End If
    Err.Clear
    On Error GoTo 0
End Function

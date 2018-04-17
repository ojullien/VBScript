' -----------------------------------------------------------------------------
' VB Scripts.
' save-to App functions
'
' @category  VB Scripts
' @package   Includes
' @version   20170905
' @copyright (©) 2017, Olivier Jullien <https://github.com/ojullien>
' -----------------------------------------------------------------------------

Option Explicit

' -----------------------------------------------------------------------------
' Robocopy
' -----------------------------------------------------------------------------

Function robocopy ( ByVal sSource, ByVal sDestination, ByVal sLogFile )
    Dim sOptionsCopy : sOptionsCopy="/Z /MIR"
    Dim sOptionsFile : sOptionsFile="/DST"
    Dim sOptionsRetry : sOptionsRetry="/R:3 /W:5"
    Dim sOptionsLog : sOptionsLog="/X /V /FP /NS /NP /TEE"
    Dim bWaitOnReturn : bWaitOnReturn = True
    Dim iWindowStyle : iWindowStyle = 0
    Dim iReturn : iReturn = -1
    Dim sCommand : sCommand = "robocopy.exe " & qq(sSource) & " " & qq(sDestination) & " " & sOptionsCopy & " " & sOptionsFile & " " & sOptionsRetry & " " & sOptionsLog & " /log:" & qq(sLogFile)
    Dim oShell
	Set oShell = CreateObject("WScript.Shell")

    On Error Resume Next
    Err.Clear

    notice "Robocopy"
    notice vbTab & "Source: " & sSource
    notice vbTab & "Destination: " & sDestination
    notice vbTab & "Log: " & sLogFile
    noticel vbTab & "Exit code: "
    iReturn = oShell.Run( sCommand, iWindowStyle, bWaitOnReturn )
    If Err.Number <> 0 Then
        error "NOK code: " & CStr(Err.Number) & " Description: " & Err.Description
        Err.Clear
        robocopy = False
    End If
    If iReturn > 8 Then
        error "NOK code: " & iReturn
        robocopy = False
    Else
        success "OK code: " & iReturn
        robocopy = True
    End If

    Err.Clear
    On Error GoTo 0

End Function

' -----------------------------------------------------------------------------
' Defragments a specified file or files
' -----------------------------------------------------------------------------

Function contigerSilent ( ByVal sPath )
    Dim sBin : sBin = "C:\Program Files\SysinternalsSuite\Contig.exe"
    Dim sOption : sOption = "-s -q -nobanner"
    Dim oShell, iErrorCode
    Dim sCommand : sCommand = qq(sBin) & " " & sOption & " " & qq(sPath & "\*")

    contigerSilent = False
    On Error Resume Next
    Err.Clear

    notice "Contiger:"
    notice vbTab & "Path: " & sPath
    noticel vbTab & "Exit returns: "

    Set oShell = WScript.CreateObject ("WScript.Shell")
    iErrorCode = oShell.run( sCommand, 0, True)
    If Err.Number <> 0 Then
        error "NOK code: " & CStr(Err.Number) & " (0x" & hex(Err.Number) & ") Description: " & Err.Description
    Else
        success "OK code: " & iErrorCode
        contigerSilent = True
    End If

    Err.Clear
    On Error GoTo 0

End Function

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

Function saveTo

    Dim iIndex : iIndex = 0
    Dim iMax : iMax = UBound(m_APP_SAVETO_LISTDIR)
    Dim sLog, sSource, sDestination, sFolder

    ' Clean log directory
    cleanDirectory m_DIR_SCRIPT & "\log"

    ' Check source
    If Not existsFolder(m_APP_SAVETO_SOURCECHK) Then
        error m_APP_SAVETO_SOURCENAME & " does not exist on " & m_APP_SAVETO_SOURCE & " ! Aborting ..."
        saveTo = False
        Exit Function
    End If

    ' Check destination
    If Not existsFolder(m_APP_SAVETO_DESTINATIONCHk) Then
        error m_APP_SAVETO_DESTINATIONNAME & " does not exist on " & m_APP_SAVETO_DESTINATION & " ! Aborting ..."
        saveTo = False
        Exit Function
    End If

    ' Ask for confirmation
    notice m_APP_SAVETO_SOURCENAME & " and " & m_APP_SAVETO_DESTINATIONNAME & " are ready."
    If Not confirmUser Then
        notice "Aborting ..."
        saveTo = False
        Exit Function
    End If

    ' Copy
    For iIndex = 0 to iMax
        sFolder = m_APP_SAVETO_LISTDIR(iIndex)
        sSource = m_APP_SAVETO_SOURCE & sFolder
        sDestination = m_APP_SAVETO_DESTINATION & sFolder
        sLog = m_DIR_LOG & "\" & WScript.ScriptName & "-" & sFolder
        separateLine
        If Not robocopy( sSource, sDestination, sLog & "-R.log" ) Then
            notice "Aborting ..."
            saveTo = False
            Exit Function
        End If
        'separateLine
        If Not contigerSilent( sDestination ) Then
            notice "Aborting ..."
            saveTo = False
            Exit Function
        End If
    Next

End Function

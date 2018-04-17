' -----------------------------------------------------
' VB Scripts.
' String functions
'
' @category  Scripts
' @package   Includes
' @version   20170911
' @copyright (©) 2017, Olivier Jullien <https://github.com/ojullien>
' -----------------------------------------------------

Option Explicit

' -----------------------------------------------------
' Find whether the type of a variable is string
' -----------------------------------------------------

Function isString( x )
    isString = False
    If vbString = VarType( x ) Then
        isString = True
    End If
End Function

Function isStringAndNotEmpty( x )
    isStringAndNotEmpty = False
    If vbString=VarType(x) Then
        ' Non determinist
         x = trim(x)
        If 0<Len(x)  Then
            isStringAndNotEmpty = True
        End If
    End If
End Function

' -----------------------------------------------------
' Write functions
' -----------------------------------------------------

Sub writeToLog ( ByVal sTxt )
    If (m_OPTION_LOG = 1) And isStringAndNotEmpty(sTxt) Then
        m_LOGFILE.WriteLine( sTxt )
    End If
End Sub

' -----------------------------------------------------
' Display functions
' -----------------------------------------------------

Sub displayError ( ByVal sTxt )
    If (m_OPTION_DISPLAY = 1) And isStringAndNotEmpty(sTxt) Then
        WScript.StdOut.WriteLine "!!! " & sTxt
    End If
End Sub

Sub displaySuccess ( ByVal sTxt )
    If (m_OPTION_DISPLAY = 1) And isStringAndNotEmpty(sTxt) Then
        WScript.StdOut.WriteLine ">>> " & sTxt
    End If
End Sub

Sub display ( ByVal sTxt )
    If (m_OPTION_DISPLAY = 1) And isStringAndNotEmpty(sTxt) Then
        WScript.StdOut.WriteLine sTxt
    End If
End Sub

Sub displayl ( ByVal sTxt )
    If (m_OPTION_DISPLAY = 1) And isStringAndNotEmpty(sTxt) Then
        WScript.StdOut.Write sTxt
    End If
End Sub

' -----------------------------------------------------
' Log functions
' -----------------------------------------------------

Sub error ( ByVal sTxt )
    If isStringAndNotEmpty(sTxt) then
        writeToLog( sTxt )
        displayError( sTxt )
    End If
End Sub

Sub notice ( ByVal sTxt )
    if isStringAndNotEmpty(sTxt) then
        writeToLog( sTxt )
        display( sTxt )
    End If
End Sub

Sub noticel ( ByVal sTxt )
    if isStringAndNotEmpty(sTxt) then
        writeToLog( sTxt )
        displayl( sTxt )
    End If
End Sub

Sub success ( ByVal sTxt )
    if isStringAndNotEmpty(sTxt) then
        writeToLog( sTxt )
        displaySuccess( sTxt )
    End If
End Sub

' -----------------------------------------------------
' Clear screen
' -----------------------------------------------------

Sub clearScreen
    If m_OPTION_DISPLAY = 1 Then
        For i = 1 To 120
            WScript.Echo
        Next
    End If
End Sub

Sub separateLine
    notice "---------------------------------------------------------------------------"
End Sub

' -----------------------------------------------------
' User interference
' -----------------------------------------------------

Function waitUser ()
    If m_OPTION_WAIT = 1 Then
        displayl "Press [ENTER] to continue."
        waitUser = LCase(WScript.StdIn.ReadLine)
    End If
    waitUser=0
End Function

Function confirmUser ()
    Dim sInput, iCompare
    displayl "Would you like to continue [y/n]?"
    sInput = LCase( Trim(WScript.StdIn.ReadLine) )
    iCompare = StrComp( "y", sInput, vbTextCompare)
    If iCompare=0 Then
        confirmUser=True
    Else
        confirmUser=False
    End If
End Function

' -----------------------------------------------------
' Double quote for command
' -----------------------------------------------------

Function qq( ByVal str )
    qq = """" & str & """"
End Function

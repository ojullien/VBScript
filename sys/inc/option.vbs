' -----------------------------------------------------
' VB Scripts.
' Options
'
' @category  VB Scripts
' @package   Includes
' @version   20170911
' @copyright (©) 2017, Olivier Jullien <https://github.com/ojullien>
' -----------------------------------------------------

Option Explicit

' -----------------------------------------------------------------------------
' log
' -----------------------------------------------------------------------------
If 1 = m_OPTION_LOG Then
    Dim oFSO
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set m_LOGFILE = oFSO.CreateTextFile( m_LOGFILEPATH, True, True)
    Set oFSO = Nothing
End If

' -----------------------------------------------------------------------------
' Trace
' -----------------------------------------------------------------------------
If 1 = m_OPTION_DISPLAY Then
    display "Display mode is ON. Contents will be displayed."

    If 1 = m_OPTION_LOG Then
        display "Log mode is ON. Contents will be logged."
    Else
        display "Log mode is OFF. Contents will not be logged."
    End If

    If 1 = m_OPTION_WAIT Then
        display "Wait mode is ON. Wait for user input between actions."
    Else
        display "Wait mode is OFF. Do not wait for user input between actions."
    End If
End If

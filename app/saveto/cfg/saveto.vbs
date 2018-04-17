' -----------------------------------------------------
' VB Scripts.
' save-to App Configuration file.
'
' @category  VB Scripts
' @package   Configuration
' @version   20170905
' @copyright (©) 2017, Olivier Jullien <https://github.com/ojullien>
' -----------------------------------------------------

Option Explicit

' -----------------------------------------------------
' Trace
' -----------------------------------------------------
separateLine
notice "App configuration"
checkDir vbTab & "Source:" & m_APP_SAVETO_SOURCENAME & " on " & m_APP_SAVETO_SOURCE& " ", m_APP_SAVETO_SOURCECHK
checkDir vbTab & "Destination:" & m_APP_SAVETO_DESTINATIONNAME & " on " & m_APP_SAVETO_DESTINATION& " ", m_APP_SAVETO_DESTINATIONCHk
Dim iIndex : iIndex = 0
Dim iMax : iMax = UBound(m_APP_SAVETO_LISTDIR)
noticel vbTab & "List directories: "
For iIndex = 0 to iMax-1
    noticel m_APP_SAVETO_LISTDIR(iIndex) & ", "
Next
notice m_APP_SAVETO_LISTDIR(iMax)

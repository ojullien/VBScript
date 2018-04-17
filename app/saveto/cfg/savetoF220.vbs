' -----------------------------------------------------
' VB Scripts.
' save-to-F220 App Configuration file.
'
' @category  VB Scripts
' @package   Configuration
' @version   20170905
' @copyright (©) 2017, Olivier Jullien <https://github.com/ojullien>
' -----------------------------------------------------

Option Explicit

' -----------------------------------------------------
' Directories
' -----------------------------------------------------

Dim m_APP_SAVETO_SOURCE : m_APP_SAVETO_SOURCE = "H:\"
Dim m_APP_SAVETO_SOURCENAME: m_APP_SAVETO_SOURCENAME = "HITACHI 250Go"
Dim m_APP_SAVETO_SOURCECHK: m_APP_SAVETO_SOURCECHK = m_APP_SAVETO_SOURCE & "250Go"

Dim m_APP_SAVETO_DESTINATION : m_APP_SAVETO_DESTINATION = "K:\"
Dim m_APP_SAVETO_DESTINATIONNAME : m_APP_SAVETO_DESTINATIONNAME = "FUJISTU 220Go"
Dim m_APP_SAVETO_DESTINATIONCHk : m_APP_SAVETO_DESTINATIONCHk = m_APP_SAVETO_DESTINATION & "220Go"

Dim m_APP_SAVETO_LISTDIR : m_APP_SAVETO_LISTDIR=Array( "Games", "Systems", "Sport" )

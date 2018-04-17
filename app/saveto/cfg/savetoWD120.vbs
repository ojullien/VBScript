' -----------------------------------------------------
' VB Scripts.
' save-to-WD120 App Configuration file.
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

Dim m_APP_SAVETO_DESTINATION : m_APP_SAVETO_DESTINATION = "J:\"
Dim m_APP_SAVETO_DESTINATIONNAME : m_APP_SAVETO_DESTINATIONNAME = "WD PASSPORT 120Go"
Dim m_APP_SAVETO_DESTINATIONCHk : m_APP_SAVETO_DESTINATIONCHk = m_APP_SAVETO_DESTINATION & "120Go"

Dim m_APP_SAVETO_LISTDIR : m_APP_SAVETO_LISTDIR=Array( "Docs", "Work", "Ebook", "Design", "Code", "Soft", "Drivers" )

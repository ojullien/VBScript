' -----------------------------------------------------
' VB Scripts.
' Main Configuration file.
'
' @category  VB Scripts
' @package   Configuration
' @version   20170905
' @copyright (©) 2017, Olivier Jullien <https://github.com/ojullien>
' -----------------------------------------------------

Option Explicit

' -----------------------------------------------------
' Date
' -----------------------------------------------------
Dim m_DATE : m_DATE = getNow()

' -----------------------------------------------------
' Main Directories
' -----------------------------------------------------
' Directory holds scripts
Dim m_DIR_SCRIPT : m_DIR_SCRIPT = pwd()
' Directory holds system files
Dim m_DIR_SYS : m_DIR_SYS = m_DIR_SCRIPT & "\sys"
' Directory holds app files
Dim m_DIR_APP : m_DIR_APP = m_DIR_SCRIPT& "\app"
' Directory holds log
Dim m_DIR_LOG : m_DIR_LOG = m_DIR_SCRIPT& "\log"

' -----------------------------------------------------
' Main Files
' -----------------------------------------------------
Dim m_LOGFILEPATH : m_LOGFILEPATH = m_DIR_LOG & "\" & m_DATE & "_" & WScript.ScriptName & ".log"
Dim m_LOGFILE

' -----------------------------------------------------
' Trace
' -----------------------------------------------------
separateLine
notice "Main configuration"
checkDir vbTab & "Script directory: "  & m_DIR_SCRIPT & " ", m_DIR_SCRIPT
checkDir vbTab & "System directory: "  & m_DIR_SYS & " ", m_DIR_SYS
checkDir vbTab & "App directory: "  & m_DIR_APP & " ", m_DIR_APP
checkDir vbTab & "Log directory: "  & m_DIR_LOG & " ", m_DIR_LOG
checkFile vbTab & "Log file is: "  & m_LOGFILEPATH & " ", m_LOGFILEPATH

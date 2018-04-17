' -----------------------------------------------------
' VB Scripts.
' unittest App Configuration file.
'
' @category  VB Scripts
' @package   Configuration
' @version   20170911
' @copyright (©) 2017, Olivier Jullien <https://github.com/ojullien>
' -----------------------------------------------------

Option Explicit

' -----------------------------------------------------
' Constants
' -----------------------------------------------------
Dim m_APP_UNITTEST_VALUES01 : m_APP_UNITTEST_VALUES01 = Array( "this is a string", "", " ", 1, 1.1, True, Null, Empty, New RegExp)
Dim m_APP_UNITTEST_VALUES02 : m_APP_UNITTEST_VALUES02 = Array( WScript.ScriptFullName, "", " ", 1, 1.1, True, Null, Empty, New RegExp)

' -----------------------------------------------------
' Trace
' -----------------------------------------------------
separateLine
notice "App configuration"

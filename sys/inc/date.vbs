' -----------------------------------------------------------------------------
' VB Scripts.
' Date functions
'
' @category  VB Scripts
' @package   Includes
' @version   20170905
' @copyright (©) 2017, Olivier Jullien <https://github.com/ojullien>
' -----------------------------------------------------------------------------

Option Explicit

' -----------------------------------------------------------------------------
' Returns the current system date and time as an expression formatted
' -----------------------------------------------------------------------------

Function getNow ()
    Dim sNow : sNow = Now
    getNow = Year(sNow) & Right( "0" & Month(sNow), 2) & Right( "0" & Day(sNow), 2) & "_" & Right( "0" & Hour(sNow), 2) & Right( "0" & Minute(sNow), 2)
End Function

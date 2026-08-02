' vybe-test: vb/vb_system_console_matrix/console_buffer_and_cursor_properties
' origin: languages/vb/tests/vb/test_vb_system_console_matrix.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Imports System

Module M
    Sub Main()
        Dim oldColor As ConsoleColor = Console.ForegroundColor
        Console.ForegroundColor = ConsoleColor.Yellow
        __Check(CStr(Console.ForegroundColor = ConsoleColor.Yellow), "True")
        Console.ForegroundColor = oldColor
    End Sub
End Module

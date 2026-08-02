' vybe-test: vb/vb_directcast_value_types/test_vb_directcast_exception_hierarchy_exact_match
' origin: languages/vb/tests/vb/test_vb_directcast_value_types.rs

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

Module Program
    Sub Main()
        Dim exObj As Object = New FormatException("Invalid Format")
        Dim fe As FormatException = DirectCast(exObj, FormatException)
        __Check(CStr(fe.Message), "Invalid Format")
    End Sub
End Module

' vybe-test: vb/vb_action_func_delegates/test_vb_func_return_value
' origin: languages/vb/tests/vb/test_vb_action_func_delegates.rs

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
        Dim fn As Func(Of Integer, Integer, Integer) = Function(a, b) a * b
        __Check(CStr(fn(6, 7)), "42")
    End Sub
End Module

' vybe-test: vb/vb_queue_stack/stack_contains_reports_membership
' origin: languages/vb/tests/vb/test_vb_queue_stack.rs

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

Imports System.Collections.Generic

Module M
    Sub Main()
        Dim s As New Stack(Of Integer)()
        s.Push(5)
        s.Push(6)
        __Check(CStr(s.Contains(6)), "True")
        __Check(CStr(s.Contains(9)), "False")
    End Sub
End Module

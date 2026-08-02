' vybe-test: vb/vb_concurrent_stack_operations/test_vb_concurrent_stack_try_peek
' origin: languages/vb/tests/vb/test_vb_concurrent_stack_operations.rs

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

Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim cs As New ConcurrentStack(Of String)()
        cs.Push("Top")
        Dim topVal As String = Nothing
        Dim ok As Boolean = cs.TryPeek(topVal)
        __Check(CStr(ok), "True")
        __Check(CStr(topVal), "Top")
        __Check(CStr(cs.Count), "1")
    End Sub
End Module

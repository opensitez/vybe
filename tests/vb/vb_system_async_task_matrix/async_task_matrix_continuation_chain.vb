' vybe-test: vb/vb_system_async_task_matrix/async_task_matrix_continuation_chain
' origin: languages/vb/tests/vb/test_vb_system_async_task_matrix.rs

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

Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim baseTask As Task(Of Integer) = Task.Run(Function() 11)
        Dim continued As Task(Of Integer) = baseTask.ContinueWith(Function(t) t.Result + 1)
        __Check(CStr(continued.Result), "12")
    End Sub
End Module

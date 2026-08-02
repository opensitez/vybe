' vybe-test: vb/vb_system_task_timeout_matrix/task_wait_with_infinite_timeout
' origin: languages/vb/tests/vb/test_vb_system_task_timeout_matrix.rs

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

Imports System.Threading
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim fast As Task(Of Integer) = Task.Run(Function()
            Thread.Sleep(10)
            Return 10
        End Function)
        __Check(CStr(fast.Wait(Timeout.Infinite)), "True")
        __Check(CStr(fast.Result), "10")
    End Sub
End Module

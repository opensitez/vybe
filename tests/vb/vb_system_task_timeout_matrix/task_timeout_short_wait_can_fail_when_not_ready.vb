' vybe-test: vb/vb_system_task_timeout_matrix/task_timeout_short_wait_can_fail_when_not_ready
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
        Dim slow As Task = Task.Run(Sub()
            Thread.Sleep(50)
        End Sub)
        __Check(CStr(slow.Wait(1)), "False")
        __Check(CStr(slow.Wait(2000)), "True")
    End Sub
End Module

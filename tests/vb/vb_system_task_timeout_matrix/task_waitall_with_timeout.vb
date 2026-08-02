' vybe-test: vb/vb_system_task_timeout_matrix/task_waitall_with_timeout
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
        Dim a As Task = Task.Run(Sub() Thread.Sleep(10))
        Dim b As Task = Task.Run(Sub() Thread.Sleep(20))
        Dim timeoutExpired As Boolean = Not Task.WaitAll(New Task() {a, b}, 1)
        Dim completed As Boolean = Task.WaitAll(New Task() {a, b}, 2000)
        __Check(CStr(timeoutExpired), "True")
        __Check(CStr(completed), "True")
    End Sub
End Module

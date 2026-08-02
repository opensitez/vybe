' vybe-test: vb/vb_task_run_exception_capture/test_vb_task_run_with_state_object_exception
' origin: languages/vb/tests/vb/test_vb_task_run_exception_capture.rs

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
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim t = Task.Factory.StartNew(Sub(state)
            Throw New Exception("State Task Error: " & state.ToString())
        End Sub, "StateValue")

        Try
            t.Wait()
        Catch ex As AggregateException
            __Check(CStr(ex.InnerException.Message), "State Task Error: StateValue")
        End Try
    End Sub
End Module

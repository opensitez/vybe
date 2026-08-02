' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_task_when_any_first_responder
' origin: languages/vb/tests/vb/test_vb_async_linq_combined_pipeline.rs

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

Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function SlowTaskAsync() As Task(Of String)
        Await Task.Delay(50)
        Return "Slow"
    End Function

    Private Async Function FastTaskAsync() As Task(Of String)
        Await Task.Yield()
        Return "Fast"
    End Function

    Sub Main()
        Dim tSlow = SlowTaskAsync()
        Dim tFast = FastTaskAsync()

        Dim winner = Task.WhenAny(tSlow, tFast)
        winner.Wait()
        __Check(CStr(winner.Result.Result), "Fast")
    End Sub
End Module

' vybe-test: vb/vb_async_task_when_all_any/test_vb_async_task_when_any
' origin: languages/vb/tests/vb/test_vb_async_task_when_all_any.rs

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

Module Program
    Async Function SlowTask() As Task(Of String)
        Await Task.Delay(50)
        Return "Slow"
    End Function

    Async Function FastTask() As Task(Of String)
        Await Task.Delay(5)
        Return "Fast"
    End Function

    Async Function RunFirstAsync() As Task
        Dim winner As Task(Of String) = Await Task.WhenAny(SlowTask(), FastTask())
        Dim val As String = Await winner
        __Check(CStr(val), "Fast")
    End Function

    Sub Main()
        RunFirstAsync().Wait()
    End Sub
End Module

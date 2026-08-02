' vybe-test: vb/vb_async_task_delay_cancellation/test_vb_async_task_when_any_first_delay_wins
' origin: languages/vb/tests/vb/test_vb_async_task_delay_cancellation.rs

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
    Private Async Function RunRaceAsync() As Task(Of String)
        Dim tFast = Task.Run(Function() As String
            Task.Delay(5).Wait()
            Return "Fast"
        End Function)
        Dim tSlow = Task.Run(Function() As String
            Task.Delay(500).Wait()
            Return "Slow"
        End Function)

        Dim winner = Await Task.WhenAny(tFast, tSlow)
        Return Await winner
    End Function

    Sub Main()
        Dim t = RunRaceAsync()
        __Check(CStr(t.Result), "Fast")
    End Sub
End Module

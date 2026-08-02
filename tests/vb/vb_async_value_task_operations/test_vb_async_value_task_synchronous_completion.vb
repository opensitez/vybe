' vybe-test: vb/vb_async_value_task_operations/test_vb_async_value_task_synchronous_completion
' origin: languages/vb/tests/vb/test_vb_async_value_task_operations.rs

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
    Function GetCachedValueAsync(cached As Boolean) As ValueTask(Of Integer)
        If cached Then
            Return New ValueTask(Of Integer)(100)
        End If
        Return New ValueTask(Of Integer)(ComputeAsync())
    End Function

    Async Function ComputeAsync() As Task(Of Integer)
        Await Task.Delay(10)
        Return 200
    End Function

    Async Function RunAsync() As Task
        Dim v1 As Integer = Await GetCachedValueAsync(True)
        Dim v2 As Integer = Await GetCachedValueAsync(False)
        __Check(CStr(v1 & ":" & v2), "100:200")
    End Function

    Sub Main()
        RunAsync().Wait()
    End Sub
End Module

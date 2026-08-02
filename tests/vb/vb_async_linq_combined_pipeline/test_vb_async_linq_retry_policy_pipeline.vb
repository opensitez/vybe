' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_retry_policy_pipeline
' origin: languages/vb/tests/vb/test_vb_async_linq_combined_pipeline.rs

Imports System
Imports System.Threading.Tasks

Module Program
    Private attempts As Integer = 0

    Private Async Function UnreliableAsync() As Task(Of String)
        Await Task.Yield()
        attempts += 1
        If attempts < 3 Then Throw New InvalidOperationException("Fail " & attempts)
        Return "Success"
    End Function

    Sub Main()
        Dim t = Task.Run(Async Function()
            For i As Integer = 1 To 5
                Try
                    Return Await UnreliableAsync()
                Catch ex As Exception
                End Try
            Next
            Return "FailedAll"
        End Function)
        t.Wait()
        Console.WriteLine(t.Result & "|Attempts=" & attempts)
    End Sub
End Module

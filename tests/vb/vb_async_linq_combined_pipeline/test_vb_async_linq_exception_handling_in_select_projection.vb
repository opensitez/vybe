' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_exception_handling_in_select_projection
' origin: languages/vb/tests/vb/test_vb_async_linq_combined_pipeline.rs

Imports System
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function SafeDivideAsync(n As Integer) As Task(Of Double)
        Await Task.Yield()
        If n = 0 Then Throw New DivideByZeroException("Cannot divide by zero")
        Return 100.0 / n
    End Function

    Sub Main()
        Dim numbers As Integer() = {10, 0, 5}
        Dim tasks = numbers.Select(Function(n) SafeDivideAsync(n)).ToArray()

        Dim results As New System.Collections.Generic.List(Of String)()
        For Each t In tasks
            Try
                t.Wait()
                results.Add(t.Result.ToString("F0"))
            Catch ex As AggregateException
                results.Add("Error")
            End Try
        Next
        Console.WriteLine(String.Join(",", results))
    End Sub
End Module

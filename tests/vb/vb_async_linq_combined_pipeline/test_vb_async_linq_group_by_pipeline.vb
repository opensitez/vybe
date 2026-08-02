' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_group_by_pipeline
' origin: languages/vb/tests/vb/test_vb_async_linq_combined_pipeline.rs

Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Class Metric
        Public Category As String
        Public Value As Integer
    End Class

    Private Async Function GetMetricsAsync() As Task(Of List(Of Metric))
        Await Task.Yield()
        Return New List(Of Metric) From {
            New Metric With {.Category = "CPU", .Value = 50},
            New Metric With {.Category = "RAM", .Value = 70},
            New Metric With {.Category = "CPU", .Value = 60}
        }
    End Function

    Sub Main()
        Dim t = GetMetricsAsync()
        t.Wait()

        Dim grouped = t.Result.GroupBy(Function(m) m.Category)
        For Each g In grouped.OrderBy(Function(g) g.Key)
            Console.WriteLine(g.Key & ":" & g.Average(Function(m) m.Value))
        Next
    End Sub
End Module

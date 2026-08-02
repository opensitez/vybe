' vybe-test: vb/vb_system_task_completion_matrix/task_completion_fail_path_reports_error
' origin: languages/vb/tests/vb/test_vb_system_task_completion_matrix.rs

Imports System
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of Integer)()
        tcs.SetException(New InvalidOperationException("bad"))

        Try
            Console.WriteLine(tcs.Task.Result)
        Catch ex As AggregateException
            Console.WriteLine(ex.InnerException.GetType().Name)
        End Try
    End Sub
End Module

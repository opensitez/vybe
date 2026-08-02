' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_cancellation_token_propagation
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

Imports System
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks

Module Program
    Private Async Function LongWorkAsync(n As Integer, ct As CancellationToken) As Task(Of Integer)
        ct.ThrowIfCancellationRequested()
        Await Task.Yield()
        Return n * 2
    End Function

    Sub Main()
        Dim cts As New CancellationTokenSource()
        cts.Cancel()

        Dim tasks = Enumerable.Range(1, 5).Select(Function(n) LongWorkAsync(n, cts.Token)).ToArray()

        Try
            Task.WaitAll(tasks)
        Catch ex As AggregateException
            __Check(CStr("AggregateException Caught on Cancelled Pipeline"), "AggregateException Caught on Cancelled Pipeline")
        End Try
    End Sub
End Module

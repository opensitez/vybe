' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_empty_source_sequence
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
    Sub Main()
        Dim emptyItems As Integer() = {}
        Dim tasks = emptyItems.Select(Async Function(n) Await Task.FromResult(n * 2)).ToArray()
        Task.WaitAll(tasks)
        __Check(CStr(tasks.Length), "0")
    End Sub
End Module

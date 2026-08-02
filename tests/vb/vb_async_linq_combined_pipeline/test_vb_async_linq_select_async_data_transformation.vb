' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_select_async_data_transformation
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
Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function FetchDataAsync(id As Integer) As Task(Of String)
        Await Task.Yield()
        Return "Item_" & id
    End Function

    Sub Main()
        Dim ids As Integer() = {1, 2, 3}
        Dim tasks = ids.Select(Function(i) FetchDataAsync(i)).ToArray()
        Task.WaitAll(tasks)

        Dim results = tasks.Select(Function(t) t.Result).ToList()
        __Check(CStr(String.Join(",", results)), "Item_1,Item_2,Item_3")
    End Sub
End Module

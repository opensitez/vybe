' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_flat_map_select_many
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

Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function GetSubItemsAsync(category As String) As Task(Of String())
        Await Task.Yield()
        Return New String() {category & "-1", category & "-2"}
    End Function

    Sub Main()
        Dim categories As String() = {"CatA", "CatB"}
        Dim tasks = categories.Select(Function(c) GetSubItemsAsync(c)).ToArray()
        Task.WaitAll(tasks)

        Dim flattened = tasks.SelectMany(Function(t) t.Result).ToList()
        __Check(CStr(String.Join(",", flattened)), "CatA-1,CatA-2,CatB-1,CatB-2")
    End Sub
End Module

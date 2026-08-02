' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_dictionary_projection
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
    Private Async Function GetKeyValuePairAsync(id As Integer) As Task(Of KeyValuePair(Of Integer, String))
        Await Task.Yield()
        Return New KeyValuePair(Of Integer, String)(id, "Code_" & id)
    End Function

    Sub Main()
        Dim ids As Integer() = {101, 102}
        Dim tasks = ids.Select(Function(i) GetKeyValuePairAsync(i)).ToArray()
        Task.WaitAll(tasks)

        Dim dict = tasks.Select(Function(t) t.Result).ToDictionary(Function(kvp) kvp.Key, Function(kvp) kvp.Value)
        __Check(CStr(dict(101) & "|" & dict(102)), "Code_101|Code_102")
    End Sub
End Module

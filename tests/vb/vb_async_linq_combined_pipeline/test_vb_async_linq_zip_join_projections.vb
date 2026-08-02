' vybe-test: vb/vb_async_linq_combined_pipeline/test_vb_async_linq_zip_join_projections
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
    Private Async Function FetchNamesAsync() As Task(Of String())
        Await Task.Yield()
        Return New String() {"Alice", "Bob"}
    End Function

    Private Async Function FetchScoresAsync() As Task(Of Integer())
        Await Task.Yield()
        Return New Integer() {90, 85}
    End Function

    Sub Main()
        Dim tNames = FetchNamesAsync()
        Dim tScores = FetchScoresAsync()
        Task.WaitAll(tNames, tScores)

        Dim zipped = tNames.Result.Zip(tScores.Result, Function(name, score) name & "=" & score)
        __Check(CStr(String.Join("|", zipped)), "Alice=90|Bob=85")
    End Sub
End Module

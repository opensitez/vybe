' vybe-test: vb/vb_system_concurrent_collections_matrix/concurrent_dictionary_try_update
' origin: languages/vb/tests/vb/test_vb_system_concurrent_collections_matrix.rs

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

Imports System.Collections.Concurrent

Module M
    Sub Main()
        Dim values As New ConcurrentDictionary(Of String, Integer)()

        values.TryAdd("state", 1)
        Dim changed As Boolean = values.TryUpdate("state", 4, 1)
        Dim failed As Boolean = values.TryUpdate("state", 9, 1)

        __Check(CStr(changed), "True")
        __Check(CStr(failed), "False")
        __Check(CStr(values("state")), "4")
    End Sub
End Module

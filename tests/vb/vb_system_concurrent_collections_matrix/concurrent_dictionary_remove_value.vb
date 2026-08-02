' vybe-test: vb/vb_system_concurrent_collections_matrix/concurrent_dictionary_remove_value
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

        values.TryAdd("temp", 5)
        Dim removed As Integer = 0
        Dim success As Boolean = values.TryRemove("temp", removed)

        __Check(CStr(success), "True")
        __Check(CStr(removed), "5")
        __Check(CStr(values.ContainsKey("temp")), "False")
    End Sub
End Module

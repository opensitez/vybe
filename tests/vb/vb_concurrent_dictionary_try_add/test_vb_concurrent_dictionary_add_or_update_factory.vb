' vybe-test: vb/vb_concurrent_dictionary_try_add/test_vb_concurrent_dictionary_add_or_update_factory
' origin: languages/vb/tests/vb/test_vb_concurrent_dictionary_try_add.rs

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

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of String, Integer)()
        ' Add initial value 10
        Dim v1 = dict.AddOrUpdate("Counter", 10, Function(key, oldVal) oldVal + 1)
        ' Update existing value
        Dim v2 = dict.AddOrUpdate("Counter", 10, Function(key, oldVal) oldVal + 1)
        __Check(CStr(v1 & "|" & v2), "10|11")
    End Sub
End Module

' vybe-test: vb/vb_concurrent_dictionary_tryadd/test_vb_concurrent_dict_try_remove
' origin: languages/vb/tests/vb/test_vb_concurrent_dictionary_tryadd.rs

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
        Dim cd As New ConcurrentDictionary(Of Integer, String)()
        cd.TryAdd(1, "One")
        Dim removedVal As String = Nothing
        Dim ok As Boolean = cd.TryRemove(1, removedVal)
        __Check(CStr(ok), "True")
        __Check(CStr(removedVal), "One")
        __Check(CStr(cd.IsEmpty), "True")
    End Sub
End Module

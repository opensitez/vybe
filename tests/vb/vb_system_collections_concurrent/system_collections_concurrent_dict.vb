' vybe-test: vb/vb_system_collections_concurrent/system_collections_concurrent_dict
' origin: languages/vb/tests/vb/test_vb_system_collections_concurrent.rs

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
        Dim cd As New ConcurrentDictionary(Of Integer, String)()
        
        cd.TryAdd(1, "One")
        cd.TryAdd(2, "Two")
        
        Dim val As String = Nothing
        If cd.TryGetValue(1, val) Then
            __Check(CStr(val), "One")
        End If
        
        ' Update
        cd.AddOrUpdate(1, "NewOne", Function(k, oldVal) "NewOne")
        __Check(CStr(cd(1)), "NewOne")
    End Sub
End Module

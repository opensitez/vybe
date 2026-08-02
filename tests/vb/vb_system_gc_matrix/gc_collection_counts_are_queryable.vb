' vybe-test: vb/vb_system_gc_matrix/gc_collection_counts_are_queryable
' origin: languages/vb/tests/vb/test_vb_system_gc_matrix.rs

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

Module M
    Sub Main()
        Dim before0 As Integer = GC.CollectionCount(0)
        Dim before1 As Integer = GC.CollectionCount(1)

        Dim arr(9_999_999) As Byte
        arr = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()

        Dim after0 As Integer = GC.CollectionCount(0)
        Dim after1 As Integer = GC.CollectionCount(1)

        __Check(CStr(after0 >= before0), "True")
        __Check(CStr(after1 >= before1), "True")
    End Sub
End Module

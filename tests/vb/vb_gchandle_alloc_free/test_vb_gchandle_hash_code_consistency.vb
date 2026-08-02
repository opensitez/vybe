' vybe-test: vb/vb_gchandle_alloc_free/test_vb_gchandle_hash_code_consistency
' origin: languages/vb/tests/vb/test_vb_gchandle_alloc_free.rs

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

Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim obj As New Object()
        Dim handle = GCHandle.Alloc(obj)
        Dim hash1 = handle.GetHashCode()
        Dim hash2 = handle.GetHashCode()
        __Check(CStr(hash1 = hash2), "True")
        handle.Free()
    End Sub
End Module

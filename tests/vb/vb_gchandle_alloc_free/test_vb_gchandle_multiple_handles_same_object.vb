' vybe-test: vb/vb_gchandle_alloc_free/test_vb_gchandle_multiple_handles_same_object
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
        Dim h1 = GCHandle.Alloc(obj)
        Dim h2 = GCHandle.Alloc(obj)
        __Check(CStr((h1 <> h2) & "|" & Object.ReferenceEquals(h1.Target, h2.Target)), "True|True")
        h1.Free()
        h2.Free()
    End Sub
End Module

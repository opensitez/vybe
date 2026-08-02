' vybe-test: vb/vb_gchandle_alloc_free/test_vb_gchandle_inequality_operator
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
        Dim obj1 As New Object()
        Dim obj2 As New Object()
        Dim h1 = GCHandle.Alloc(obj1)
        Dim h2 = GCHandle.Alloc(obj2)
        __Check(CStr(h1 <> h2), "True")
        h1.Free()
        h2.Free()
    End Sub
End Module

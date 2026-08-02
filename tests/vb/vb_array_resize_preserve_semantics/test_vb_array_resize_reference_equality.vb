' vybe-test: vb/vb_array_resize_preserve_semantics/test_vb_array_resize_reference_equality
' origin: languages/vb/tests/vb/test_vb_array_resize_preserve_semantics.rs

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

Module Program
    Sub Main()
        Dim original As Integer() = {1, 2, 3}
        Dim reference As Integer() = original
        Array.Resize(original, 5)
        __Check(CStr(original.Length), "5")
        __Check(CStr(reference.Length), "3")
        __Check(CStr(Object.ReferenceEquals(original, reference)), "False")
    End Sub
End Module

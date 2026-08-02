' vybe-test: vb/vb_unsafe_pointer_arithmetic/test_vb_pointer_unmanaged_type_size_computations
' origin: languages/vb/tests/vb/test_vb_unsafe_pointer_arithmetic.rs

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
        Dim charSize = Marshal.SystemDefaultCharSize
        __Check(CStr(charSize = 1 OrElse charSize = 2), "True")
    End Sub
End Module

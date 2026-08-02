' vybe-test: vb/vb_unsafe_pointer_arithmetic/test_vb_pointer_aligned_offset_checks
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

Imports System

Module Program
    Sub Main()
        Dim p As New IntPtr(1024)
        Dim isAligned4 = (p.ToInt64() Mod 4L = 0)
        Dim isAligned8 = (p.ToInt64() Mod 8L = 0)
        __Check(CStr(isAligned4 & "|" & isAligned8), "True|True")
    End Sub
End Module

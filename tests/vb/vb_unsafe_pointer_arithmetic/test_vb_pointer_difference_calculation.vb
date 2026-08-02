' vybe-test: vb/vb_unsafe_pointer_arithmetic/test_vb_pointer_difference_calculation
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
        Dim p1 As New IntPtr(2000)
        Dim p2 As New IntPtr(1000)
        Dim diff As Long = p1.ToInt64() - p2.ToInt64()
        __Check(CStr(diff), "1000")
    End Sub
End Module

' vybe-test: vb/vb_intptr_uintptr_operations/test_vb_intptr_equality_and_inequality_operators
' origin: languages/vb/tests/vb/test_vb_intptr_uintptr_operations.rs

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
        Dim p1 As New IntPtr(1234)
        Dim p2 As New IntPtr(1234)
        Dim p3 As New IntPtr(5678)
        __Check(CStr((p1 = p2) & "|" & (p1 <> p3)), "True|True")
    End Sub
End Module

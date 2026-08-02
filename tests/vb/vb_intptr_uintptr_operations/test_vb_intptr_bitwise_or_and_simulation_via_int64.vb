' vybe-test: vb/vb_intptr_uintptr_operations/test_vb_intptr_bitwise_or_and_simulation_via_int64
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
        Dim p1 As New IntPtr(&HFF00)
        Dim p2 As New IntPtr(&H00FF)
        Dim combined As New IntPtr(p1.ToInt64() Or p2.ToInt64())
        __Check(CStr(Hex(combined.ToInt64())), "FFFF")
    End Sub
End Module

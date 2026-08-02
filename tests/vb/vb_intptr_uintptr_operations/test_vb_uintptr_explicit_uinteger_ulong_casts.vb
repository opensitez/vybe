' vybe-test: vb/vb_intptr_uintptr_operations/test_vb_uintptr_explicit_uinteger_ulong_casts
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
        Dim val As ULong = 8192UL
        Dim ptr As New UIntPtr(val)
        Dim backToUInt64 As ULong = ptr.ToUInt64()
        __Check(CStr(backToUInt64), "8192")
    End Sub
End Module

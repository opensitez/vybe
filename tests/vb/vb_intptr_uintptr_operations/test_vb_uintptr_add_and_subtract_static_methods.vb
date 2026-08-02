' vybe-test: vb/vb_intptr_uintptr_operations/test_vb_uintptr_add_and_subtract_static_methods
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
        Dim basePtr As New UIntPtr(2000UI)
        Dim added = UIntPtr.Add(basePtr, 100)
        Dim subtracted = UIntPtr.Subtract(added, 50)
        __Check(CStr(added.ToUInt32() & "|" & subtracted.ToUInt32()), "2100|2050")
    End Sub
End Module

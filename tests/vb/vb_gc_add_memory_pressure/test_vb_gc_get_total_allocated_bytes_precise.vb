' vybe-test: vb/vb_gc_add_memory_pressure/test_vb_gc_get_total_allocated_bytes_precise
' origin: languages/vb/tests/vb/test_vb_gc_add_memory_pressure.rs

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
        Dim b1 = GC.GetTotalAllocatedBytes(precise:=True)
        Dim dummy As New Byte(1000) {}
        Dim b2 = GC.GetTotalAllocatedBytes(precise:=True)
        __Check(CStr(b2 > b1), "True")
    End Sub
End Module

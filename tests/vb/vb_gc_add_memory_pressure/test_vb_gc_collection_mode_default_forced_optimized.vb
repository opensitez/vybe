' vybe-test: vb/vb_gc_add_memory_pressure/test_vb_gc_collection_mode_default_forced_optimized
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
        GC.Collect(0, GCCollectionMode.Default)
        GC.Collect(0, GCCollectionMode.Forced)
        GC.Collect(0, GCCollectionMode.Optimized)
        __Check(CStr("All Collection Modes Succeeded"), "All Collection Modes Succeeded")
    End Sub
End Module

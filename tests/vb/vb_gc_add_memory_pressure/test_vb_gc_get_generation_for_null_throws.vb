' vybe-test: vb/vb_gc_add_memory_pressure/test_vb_gc_get_generation_for_null_throws
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
        Try
            Dim obj As Object = Nothing
            GC.GetGeneration(obj)
        Catch ex As ArgumentNullException
            __Check(CStr("ArgumentNullException Caught on Null GetGeneration"), "ArgumentNullException Caught on Null GetGeneration")
        End Try
    End Sub
End Module

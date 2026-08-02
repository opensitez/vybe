' vybe-test: vb/vb_finalizer_suppress_finalize/test_vb_finalizer_large_object_heap_finalization
' origin: languages/vb/tests/vb/test_vb_finalizer_suppress_finalize.rs

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

Class LargeFinalizableObject
    Private buffer(100000) As Byte ' Placed on Large Object Heap (LOH)
    Public Shared Finalized As Boolean = False

    Protected Overrides Sub Finalize()
        Finalized = True
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Dim l As New LargeFinalizableObject()
        End Sub()

        GC.Collect(2, GCCollectionMode.Forced)
        GC.WaitForPendingFinalizers()

        __Check(CStr(LargeFinalizableObject.Finalized), "True")
    End Sub
End Module

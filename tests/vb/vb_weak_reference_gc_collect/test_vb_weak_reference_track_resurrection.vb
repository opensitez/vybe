' vybe-test: vb/vb_weak_reference_gc_collect/test_vb_weak_reference_track_resurrection
' origin: languages/vb/tests/vb/test_vb_weak_reference_gc_collect.rs

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

Class ResurrectedObject
    Public Shared Holder As ResurrectedObject
    Protected Overrides Sub Finalize()
        Holder = Me ' Resurrect object!
    End Sub
End Class

Module Program
    Sub Main()
        Dim weakRefShort As WeakReference
        Dim weakRefLong As WeakReference

        Sub()
            Dim obj As New ResurrectedObject()
            weakRefShort = New WeakReference(obj, trackResurrection:=False)
            weakRefLong = New WeakReference(obj, trackResurrection:=True)
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()

        __Check(CStr("LongTrackAlive: " & (ResurrectedObject.Holder IsNot Nothing)), "LongTrackAlive: True")
    End Sub
End Module

' vybe-test: vb/vb_gc_add_memory_pressure/test_vb_gc_add_and_remove_memory_pressure
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

Class NativeBufferHolder
    Private bytesAllocated As Long

    Public Sub New(size As Long)
        bytesAllocated = size
        GC.AddMemoryPressure(bytesAllocated)
    End Sub

    Public Sub Release()
        If bytesAllocated > 0 Then
            GC.RemoveMemoryPressure(bytesAllocated)
            bytesAllocated = 0
        End If
    End Sub
End Class

Module Program
    Sub Main()
        Dim holder As New NativeBufferHolder(1024 * 1024 * 10) ' 10MB pressure
        __Check(CStr("Added Pressure"), "Added Pressure")
        holder.Release()
        __Check(CStr("Removed Pressure"), "Removed Pressure")
    End Sub
End Module

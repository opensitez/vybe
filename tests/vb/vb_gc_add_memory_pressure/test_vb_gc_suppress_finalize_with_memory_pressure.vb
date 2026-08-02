' vybe-test: vb/vb_gc_add_memory_pressure/test_vb_gc_suppress_finalize_with_memory_pressure
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

Class NativeResource
    Implements IDisposable
    Private allocatedBytes As Long

    Public Sub New(bytes As Long)
        allocatedBytes = bytes
        GC.AddMemoryPressure(allocatedBytes)
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        If allocatedBytes > 0 Then
            GC.RemoveMemoryPressure(allocatedBytes)
            allocatedBytes = 0
        End If
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overrides Sub Finalize()
        Dispose()
    End Sub
End Class

Module Program
    Sub Main()
        Using res As New NativeResource(5000000)
            __Check(CStr("Native Resource Active"), "Native Resource Active")
        End Using
    End Sub
End Module

' vybe-test: vb/vb_gchandle_alloc_free/test_vb_gchandle_to_intptr_and_from_intptr_roundtrip
' origin: languages/vb/tests/vb/test_vb_gchandle_alloc_free.rs

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
Imports System.Runtime.InteropServices

Class Container
    Public Title As String = "VybeHandle"
End Class

Module Program
    Sub Main()
        Dim c As New Container()
        Dim handle = GCHandle.Alloc(c)
        Dim ptr As IntPtr = GCHandle.ToIntPtr(handle)

        Dim restoredHandle = GCHandle.FromIntPtr(ptr)
        Dim restored As Container = CType(restoredHandle.Target, Container)
        __Check(CStr(restored.Title), "VybeHandle")
        handle.Free()
    End Sub
End Module

' vybe-test: vb/vb_gchandle_alloc_free/test_vb_gchandle_pinned_blittable_struct
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

<StructLayout(LayoutKind.Sequential)>
Structure Vector3
    Public X As Single
    Public Y As Single
    Public Z As Single
End Structure

Module Program
    Sub Main()
        Dim v As New Vector3 With {.X = 1.0F, .Y = 2.0F, .Z = 3.0F}
        Dim handle = GCHandle.Alloc(v, GCHandleType.Pinned)
        Dim addr = handle.AddrOfPinnedObject()
        __Check(CStr(addr <> IntPtr.Zero), "True")
        handle.Free()
    End Sub
End Module

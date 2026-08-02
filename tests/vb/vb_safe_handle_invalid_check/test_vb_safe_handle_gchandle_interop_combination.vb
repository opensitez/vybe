' vybe-test: vb/vb_safe_handle_invalid_check/test_vb_safe_handle_gchandle_interop_combination
' origin: languages/vb/tests/vb/test_vb_safe_handle_invalid_check.rs

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

Module Program
    Sub Main()
        Dim data = "InteropPayload"
        Dim gcHandle = GCHandle.Alloc(data, GCHandleType.Pinned)
        Dim ptr = gcHandle.AddrOfPinnedObject()
        __Check(CStr(ptr <> IntPtr.Zero), "True")
        gcHandle.Free()
    End Sub
End Module

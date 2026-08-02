' vybe-test: vb/vb_marshal_size_of_structure/test_vb_marshal_string_to_hglobal_ansi_and_free
' origin: languages/vb/tests/vb/test_vb_marshal_size_of_structure.rs

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
        Dim ptr As IntPtr = Marshal.StringToHGlobalAnsi("NativeAnsi")
        Dim restored = Marshal.PtrToStringAnsi(ptr)
        Marshal.FreeHGlobal(ptr)
        __Check(CStr(restored), "NativeAnsi")
    End Sub
End Module

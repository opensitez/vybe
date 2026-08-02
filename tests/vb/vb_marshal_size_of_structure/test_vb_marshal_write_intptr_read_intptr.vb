' vybe-test: vb/vb_marshal_size_of_structure/test_vb_marshal_write_intptr_read_intptr
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
        Dim targetPtr As New IntPtr(1024)
        Dim buffer As IntPtr = Marshal.AllocHGlobal(IntPtr.Size)
        Marshal.WriteIntPtr(buffer, targetPtr)
        Dim readPtr = Marshal.ReadIntPtr(buffer)
        Marshal.FreeHGlobal(buffer)
        __Check(CStr(readPtr.ToInt64()), "1024")
    End Sub
End Module

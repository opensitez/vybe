' vybe-test: vb/vb_marshal_size_of_structure/test_vb_marshal_get_function_pointer_for_delegate
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

Delegate Function BinaryOp(a As Integer, b As Integer) As Integer

Module Program
    Private Function Add(a As Integer, b As Integer) As Integer
        Return a + b
    End Function

    Sub Main()
        Dim del As BinaryOp = AddressOf Add
        Dim funcPtr As IntPtr = Marshal.GetFunctionPointerForDelegate(del)
        __Check(CStr(funcPtr <> IntPtr.Zero), "True")
    End Sub
End Module

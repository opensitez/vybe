' vybe-test: vb/vb_marshal_size_of_structure/test_vb_marshal_get_delegate_for_function_pointer
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

Delegate Function ComputeFunc(x As Integer) As Integer

Module Program
    Private Function Square(x As Integer) As Integer
        Return x * x
    End Function

    Sub Main()
        Dim delOrig As ComputeFunc = AddressOf Square
        Dim ptr = Marshal.GetFunctionPointerForDelegate(delOrig)
        Dim delRestored As ComputeFunc = CType(Marshal.GetDelegateForFunctionPointer(ptr, GetType(ComputeFunc)), ComputeFunc)
        __Check(CStr(delRestored(5)), "25")
    End Sub
End Module

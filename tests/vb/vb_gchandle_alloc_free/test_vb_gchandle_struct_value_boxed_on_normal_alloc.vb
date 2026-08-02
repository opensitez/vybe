' vybe-test: vb/vb_gchandle_alloc_free/test_vb_gchandle_struct_value_boxed_on_normal_alloc
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

Imports System.Runtime.InteropServices

Structure ValueHolder
    Public Count As Integer
End Structure

Module Program
    Sub Main()
        Dim v As New ValueHolder With {.Count = 99}
        Dim handle = GCHandle.Alloc(v) ' Boxing occurs for value types!
        Dim boxed As ValueHolder = CType(handle.Target, ValueHolder)
        __Check(CStr(boxed.Count), "99")
        handle.Free()
    End Sub
End Module

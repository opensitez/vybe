' vybe-test: vb/vb_generic_struct_methods/test_vb_generic_struct_unboxing_cast
' origin: languages/vb/tests/vb/test_vb_generic_struct_methods.rs

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

Structure SimpleBox(Of T)
    Public Value As T
    Public Sub New(v As T)
        Value = v
    End Sub
End Structure

Module Program
    Sub Main()
        Dim box As Object = New SimpleBox(Of Integer)(42)
        Dim unboxed = CType(box, SimpleBox(Of Integer))
        __Check(CStr(unboxed.Value), "42")
    End Sub
End Module

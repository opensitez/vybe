' vybe-test: vb/vb_ctype_custom_operator/test_vb_ctype_generic_struct_conversion_operator
' origin: languages/vb/tests/vb/test_vb_ctype_custom_operator.rs

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

Structure Wrapper(Of T)
    Public Value As T
    Public Sub New(v As T)
        Value = v
    End Sub

    Public Shared Widening Operator CType(v As T) As Wrapper(Of T)
        Return New Wrapper(Of T)(v)
    End Shared Widening Operator
End Structure

Module Program
    Sub Main()
        Dim w As Wrapper(Of String) = CType("WrappedString", Wrapper(Of String))
        __Check(CStr(w.Value), "WrappedString")
    End Sub
End Module

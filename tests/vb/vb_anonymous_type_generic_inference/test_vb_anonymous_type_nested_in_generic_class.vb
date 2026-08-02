' vybe-test: vb/vb_anonymous_type_generic_inference/test_vb_anonymous_type_nested_in_generic_class
' origin: languages/vb/tests/vb/test_vb_anonymous_type_generic_inference.rs

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

Class Wrapper(Of T)
    Public Function Wrap(val As T) As Object
        Return New With {.WrappedVal = val}
    End Function
End Class

Module Program
    Sub Main()
        Dim w As New Wrapper(Of Double)()
        Dim anon As Dynamic = w.Wrap(3.14159)
        __Check(CStr(anon.WrappedVal), "3.14159")
    End Sub
End Module

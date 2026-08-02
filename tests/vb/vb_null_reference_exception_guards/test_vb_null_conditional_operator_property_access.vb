' vybe-test: vb/vb_null_reference_exception_guards/test_vb_null_conditional_operator_property_access
' origin: languages/vb/tests/vb/test_vb_null_reference_exception_guards.rs

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

Class Address
    Public Property City As String = "Seattle"
End Class

Class Person
    Public Property HomeAddress As Address
End Class

Module Program
    Sub Main()
        Dim p As New Person()
        __Check(CStr(p.HomeAddress?.City Is Nothing), "True")
        p.HomeAddress = New Address()
        __Check(CStr(p.HomeAddress?.City), "Seattle")
    End Sub
End Module

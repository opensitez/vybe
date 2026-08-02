' vybe-test: vb/vb_property_protected_internal_mix/test_vb_property_write_only_private_backing_field
' origin: languages/vb/tests/vb/test_vb_property_protected_internal_mix.rs

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

Class SystemToken
    Private _token As String
    Public WriteOnly Property Token As String
        Set(value As String)
            _token = "ENC_" & value
        End Set
    End Property
    Public Function ValidateToken(input As String) As Boolean
        Return _token = "ENC_" & input
    End Function
End Class

Module Program
    Sub Main()
        Dim st As New SystemToken()
        st.Token = "Pass123"
        __Check(CStr(st.ValidateToken("Pass123")), "True")
    End Sub
End Module

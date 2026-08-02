' vybe-test: vb/vb_property_protected_internal_mix/test_vb_property_protected_getter_and_setter
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

Class BaseService
    Protected Property SecretKey As String = "BaseKey"
End Class

Class CustomService
    Inherits BaseService
    Public Function GetKey() As String
        Return SecretKey
    End Function
End Class

Module Program
    Sub Main()
        Dim s As New CustomService()
        __Check(CStr(s.GetKey()), "BaseKey")
    End Sub
End Module

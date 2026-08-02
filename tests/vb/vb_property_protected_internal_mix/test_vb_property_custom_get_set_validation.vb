' vybe-test: vb/vb_property_protected_internal_mix/test_vb_property_custom_get_set_validation
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

Class AgeValidator
    Private _age As Integer
    Public Property Age As Integer
        Get
            Return _age
        End Get
        Set(value As Integer)
            If value >= 0 Then _age = value Else _age = 0
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim a As New AgeValidator()
        a.Age = 25
        __Check(CStr(a.Age), "25")
        a.Age = -5
        __Check(CStr(a.Age), "0")
    End Sub
End Module

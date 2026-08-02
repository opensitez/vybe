' vybe-test: vb/vb_property_protected_internal_mix/test_vb_property_read_only_with_backing_field_private
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

Class Product
    Private _price As Decimal
    Public ReadOnly Property Price As Decimal
        Get
            Return _price
        End Get
    End Property
    Public Sub New(p As Decimal)
        _price = p
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New Product(29.99D)
        __Check(CStr(p.Price), "29.99")
    End Sub
End Module

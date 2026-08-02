' vybe-test: vb/vb_full_domain_model_simulation/test_vb_domain_coupon_code_validation
' origin: languages/vb/tests/vb/test_vb_full_domain_model_simulation.rs

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

Class Coupon
    Public Property Code As String
    Public Property ExpiryYear As Integer
    Public Property MinOrderAmount As Decimal

    Public Function IsValid(year As Integer, total As Decimal) As Boolean
        Return year <= ExpiryYear AndAlso total >= MinOrderAmount
    End Function
End Class

Module Program
    Sub Main()
        Dim c As New Coupon With {.Code = "SAVE20", .ExpiryYear = 2026, .MinOrderAmount = 50D}
        __Check(CStr(c.IsValid(2025, 75D) & "|" & c.IsValid(2025, 30D)), "True|False")
    End Sub
End Module

' vybe-test: vb/vb_full_domain_model_simulation/test_vb_domain_value_object_address_equality
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

Imports System

Class Address
    Implements IEquatable(Of Address)

    Public Property Street As String
    Public Property City As String
    Public Property Zip As String

    Public Function Equals1(other As Address) As Boolean Implements IEquatable(Of Address).Equals
        If other Is Nothing Then Return False
        Return Street = other.Street AndAlso City = other.City AndAlso Zip = other.Zip
    End Function
End Class

Module Program
    Sub Main()
        Dim a1 As New Address With {.Street = "123 Main St", .City = "NY", .Zip = "10001"}
        Dim a2 As New Address With {.Street = "123 Main St", .City = "NY", .Zip = "10001"}
        __Check(CStr(a1.Equals1(a2)), "True")
    End Sub
End Module

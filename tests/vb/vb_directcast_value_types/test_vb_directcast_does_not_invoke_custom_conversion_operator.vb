' vybe-test: vb/vb_directcast_value_types/test_vb_directcast_does_not_invoke_custom_conversion_operator
' origin: languages/vb/tests/vb/test_vb_directcast_value_types.rs

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

Class Money
    Public Amount As Decimal
    Public Sub New(a As Decimal)
        Amount = a
    End Sub

    Public Shared Widening Operator CType(a As Decimal) As Money
        Return New Money(a)
    End Shared Widening Operator
End Class

Module Program
    Sub Main()
        Dim boxed As Object = 99.9D
        Try
            ' DirectCast does not call user-defined CType conversion operator!
            Dim m As Money = DirectCast(boxed, Money)
        Catch ex As InvalidCastException
            __Check(CStr("InvalidCastException Caught on Custom Operator DirectCast"), "InvalidCastException Caught on Custom Operator DirectCast")
        End Try
    End Sub
End Module

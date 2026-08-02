' vybe-test: vb/vb_ctype_custom_operator/test_vb_ctype_bidirectional_conversion_operators
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

Class Celsius
    Public Degrees As Double
    Public Sub New(d As Double)
        Degrees = d
    End Sub
End Class

Class Fahrenheit
    Public Degrees As Double
    Public Sub New(d As Double)
        Degrees = d
    End Sub

    Public Shared Widening Operator CType(c As Celsius) As Fahrenheit
        Return New Fahrenheit(c.Degrees * 9.0 / 5.0 + 32.0)
    End Shared Widening Operator

    Public Shared Widening Operator CType(f As Fahrenheit) As Celsius
        Return New Celsius((f.Degrees - 32.0) * 5.0 / 9.0)
    End Shared Widening Operator
End Class

Module Program
    Sub Main()
        Dim c As New Celsius(100)
        Dim f As Fahrenheit = CType(c, Fahrenheit)
        Dim restoredC As Celsius = CType(f, Celsius)
        __Check(CStr(f.Degrees & "|" & restoredC.Degrees), "212|100")
    End Sub
End Module

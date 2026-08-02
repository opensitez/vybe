' vybe-test: vb/vb_ctype_custom_operator/test_vb_ctype_chained_custom_conversions
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

Class Meter
    Public Value As Double
    Public Sub New(v As Double)
        Value = v
    End Sub
    Public Shared Widening Operator CType(v As Double) As Meter
        Return New Meter(v)
    End Shared Widening Operator
End Class

Class Kilometer
    Public Value As Double
    Public Sub New(v As Double)
        Value = v
    End Sub
    Public Shared Widening Operator CType(m As Meter) As Kilometer
        Return New Kilometer(m.Value / 1000.0)
    End Shared Widening Operator
End Class

Module Program
    Sub Main()
        Dim m As Meter = CType(2500.0, Meter)
        Dim km As Kilometer = CType(m, Kilometer)
        __Check(CStr(km.Value), "2.5")
    End Sub
End Module

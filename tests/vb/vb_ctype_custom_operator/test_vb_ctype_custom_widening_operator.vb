' vybe-test: vb/vb_ctype_custom_operator/test_vb_ctype_custom_widening_operator
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

Imports System

Class Distance
    Public Meters As Double
    Public Sub New(m As Double)
        Meters = m
    End Sub

    ' Widening operator: Double to Distance
    Public Shared Widening Operator CType(m As Double) As Distance
        Return New Distance(m)
    End Shared Widening Operator
End Class

Module Program
    Sub Main()
        ' Explicit or implicit CType call uses Widening operator!
        Dim d As Distance = CType(100.5, Distance)
        __Check(CStr(d.Meters), "100.5")
    End Sub
End Module

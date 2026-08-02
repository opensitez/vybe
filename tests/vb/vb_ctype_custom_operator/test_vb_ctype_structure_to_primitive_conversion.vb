' vybe-test: vb/vb_ctype_custom_operator/test_vb_ctype_structure_to_primitive_conversion
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

Structure ComplexNumber
    Public Real As Double
    Public Imaginary As Double
    Public Sub New(r As Double, i As Double)
        Real = r
        Imaginary = i
    End Sub

    Public Shared Narrowing Operator CType(c As ComplexNumber) As Double
        Return c.Real
    End Shared Narrowing Operator
End Structure

Module Program
    Sub Main()
        Dim c As New ComplexNumber(42.5, 3.0)
        Dim r As Double = CType(c, Double)
        __Check(CStr(r), "42.5")
    End Sub
End Module

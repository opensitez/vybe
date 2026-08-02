' vybe-test: vb/vb_property_protected_internal_mix/test_vb_property_expression_bodied_public_get
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

Class Circle
    Public Property Radius As Double
    Public ReadOnly Property Area As Double => Math.PI * Radius * Radius
    Public Sub New(r As Double)
        Radius = r
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As New Circle(10)
        __Check(CStr(Math.Round(c.Area, 2)), "314.16")
    End Sub
End Module

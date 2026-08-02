' vybe-test: vb/vb_oop_polymorphism/oop_mustinherit_mustoverride
' origin: languages/vb/tests/vb/test_vb_oop_polymorphism.rs

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

MustInherit Class Shape
    Public MustOverride Function Area() As Double
End Class

Class Circle
    Inherits Shape
    Public Radius As Double
    
    Public Overrides Function Area() As Double
        Return 3.14 * Radius * Radius
    End Function
End Class

Module M
    Sub Main()
        Dim c As New Circle() With {.Radius = 10}
        Dim s As Shape = c
        __Check(CStr(s.Area()), "314")
    End Sub
End Module

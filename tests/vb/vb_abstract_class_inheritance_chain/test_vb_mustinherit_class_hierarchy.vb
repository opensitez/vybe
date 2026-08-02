' vybe-test: vb/vb_abstract_class_inheritance_chain/test_vb_mustinherit_class_hierarchy
' origin: languages/vb/tests/vb/test_vb_abstract_class_inheritance_chain.rs

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
    Public MustOverride Function GetArea() As Double
End Class

MustInherit Class Polygon
    Inherits Shape
    Public MustOverride Function GetSides() As Integer
End Class

Class Rectangle
    Inherits Polygon
    Public Width As Double = 5.0
    Public Height As Double = 4.0

    Public Overrides Function GetArea() As Double
        Return Width * Height
    End Function

    Public Overrides Function GetSides() As Integer
        Return 4
    End Function
End Class

Module Program
    Sub Main()
        Dim rect As Shape = New Rectangle()
        __Check(CStr(rect.GetArea()), "20")
        Dim poly As Polygon = CType(rect, Polygon)
        __Check(CStr(poly.GetSides()), "4")
    End Sub
End Module

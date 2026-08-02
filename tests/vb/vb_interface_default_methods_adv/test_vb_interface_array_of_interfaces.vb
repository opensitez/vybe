' vybe-test: vb/vb_interface_default_methods_adv/test_vb_interface_array_of_interfaces
' origin: languages/vb/tests/vb/test_vb_interface_default_methods_adv.rs

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

Interface IShape
    Function Area() As Double
End Interface

Class Circle
    Implements IShape
    Public Radius As Double
    Public Sub New(r As Double) : Radius = r : End Sub
    Public Function Area() As Double Implements IShape.Area
        Return Math.PI * Radius * Radius
    End Function
End Class

Class Square
    Implements IShape
    Public Side As Double
    Public Sub New(s As Double) : Side = s : End Sub
    Public Function Area() As Double Implements IShape.Area
        Return Side * Side
    End Function
End Class

Module Program
    Sub Main()
        Dim shapes As IShape() = {New Circle(10), New Square(10)}
        __Check(CStr(Math.Round(shapes(0).Area(), 2) & "|" & shapes(1).Area()), "314.16|100")
    End Sub
End Module

' vybe-test: vb/vb_mustoverride_overrides/mustoverride_and_overrides
' origin: languages/vb/tests/vb/test_vb_mustoverride_overrides.rs

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
    Public MustOverride Property Name As String
End Class

Class Circle
    Inherits Shape
    
    Private _name As String = "Circle"
    Private _radius As Double
    
    Public Sub New(radius As Double)
        _radius = radius
    End Sub
    
    Public Overrides Function GetArea() As Double
        Return Math.PI * _radius * _radius
    End Function
    
    Public Overrides Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim s As Shape = New Circle(10)
        __Check(CStr(s.Name), "Circle")
        __Check(CStr(Math.Round(s.GetArea())), "314")
    End Sub
End Module

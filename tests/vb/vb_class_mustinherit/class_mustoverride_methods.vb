' vybe-test: vb/vb_class_mustinherit/class_mustoverride_methods
' origin: languages/vb/tests/vb/test_vb_class_mustinherit.rs

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
    Public MustOverride Function Area() As Integer
End Class

Class Square
    Inherits Shape
    
    Private _side As Integer
    Public Sub New(side As Integer)
        _side = side
    End Sub
    
    Public Overrides Function Area() As Integer
        Return _side * _side
    End Function
End Class

Module M
    Sub Main()
        Dim s As Shape = New Square(4)
        __Check(CStr(s.Area()), "16")
    End Sub
End Module

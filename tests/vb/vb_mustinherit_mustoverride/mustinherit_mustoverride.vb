' vybe-test: vb/vb_mustinherit_mustoverride/mustinherit_mustoverride
' origin: languages/vb/tests/vb/test_vb_mustinherit_mustoverride.rs

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
    
    Public Sub Print()
        __Check(CStr("Area: " & GetArea()), "Area: 25")
    End Sub
End Class

Class Square
    Inherits Shape
    
    Public Property Side As Double
    
    Public Overrides Function GetArea() As Double
        Return Side * Side
    End Function
End Class

Module M
    Sub Main()
        Dim s As Shape = New Square() With {.Side = 5}
        s.Print()
    End Sub
End Module

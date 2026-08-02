' vybe-test: vb/vb_interfaces_explicit/interface_explicit_implementation_different_name
' origin: languages/vb/tests/vb/test_vb_interfaces_explicit.rs

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
    Sub Draw()
End Interface

Class Circle
    Implements IShape
    
    ' Explicitly implementing with a different method name
    Private Sub Render() Implements IShape.Draw
        __Check(CStr("Drawing Circle"), "This is class Draw, not interface Draw")
    End Sub
    
    Public Sub Draw()
        __Check(CStr("This is class Draw, not interface Draw"), "Drawing Circle")
    End Sub
End Class

Module M
    Sub Main()
        Dim c As New Circle()
        c.Draw() ' Calls class method
        
        Dim s As IShape = c
        s.Draw() ' Calls interface method (Render)
    End Sub
End Module

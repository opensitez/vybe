' vybe-test: vb/vb_operator_overloading_adv/operator_overloading_unary
' origin: languages/vb/tests/vb/test_vb_operator_overloading_adv.rs

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

Class Vector
    Public X, Y As Integer
    
    Public Sub New(x As Integer, y As Integer)
        Me.X = x
        Me.Y = y
    End Sub
    
    ' Unary operator -
    Public Shared Operator -(v As Vector) As Vector
        Return New Vector(-v.X, -v.Y)
    End Operator
    
    ' Unary operator Not
    Public Shared Operator Not(v As Vector) As Vector
        Return New Vector(Not v.X, Not v.Y)
    End Operator
End Class

Module M
    Sub Main()
        Dim v As New Vector(5, -10)
        Dim vNeg = -v
        __Check(CStr(vNeg.X), "-5")
        __Check(CStr(vNeg.Y), "10")
        
        Dim vNot = Not v
        __Check(CStr(vNot.X), "-6") ' Not 5 = -6
    End Sub
End Module

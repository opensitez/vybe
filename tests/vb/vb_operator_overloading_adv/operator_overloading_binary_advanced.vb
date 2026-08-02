' vybe-test: vb/vb_operator_overloading_adv/operator_overloading_binary_advanced
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
    
    ' Binary operator * (scalar multiplication)
    Public Shared Operator *(v As Vector, scalar As Integer) As Vector
        Return New Vector(v.X * scalar, v.Y * scalar)
    End Operator
    
    ' Binary operator * (scalar multiplication reversed)
    Public Shared Operator *(scalar As Integer, v As Vector) As Vector
        Return New Vector(v.X * scalar, v.Y * scalar)
    End Operator
End Class

Module M
    Sub Main()
        Dim v As New Vector(2, 3)
        Dim v1 = v * 5
        Dim v2 = 10 * v
        
        __Check(CStr(v1.X), "10")
        __Check(CStr(v2.Y), "30")
    End Sub
End Module

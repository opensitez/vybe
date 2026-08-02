' vybe-test: vb/vb_operator_overloading_custom/operator_overloading_custom
' origin: languages/vb/tests/vb/test_vb_operator_overloading_custom.rs

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
    Public X As Integer
    Public Y As Integer
    
    Public Shared Operator +(v1 As Vector, v2 As Vector) As Vector
        Return New Vector() With {.X = v1.X + v2.X, .Y = v1.Y + v2.Y}
    End Operator
    
    Public Shared Operator -(v1 As Vector) As Vector
        Return New Vector() With {.X = -v1.X, .Y = -v1.Y}
    End Operator
End Class

Module M
    Sub Main()
        Dim v1 As New Vector() With {.X = 1, .Y = 2}
        Dim v2 As New Vector() With {.X = 3, .Y = 4}
        
        Dim v3 = v1 + v2
        Dim v4 = -v1
        
        __Check(CStr(v3.X), "4")
        __Check(CStr(v4.X), "-1")
    End Sub
End Module

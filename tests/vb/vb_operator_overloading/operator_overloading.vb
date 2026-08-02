' vybe-test: vb/vb_operator_overloading/operator_overloading
' origin: languages/vb/tests/vb/test_vb_operator_overloading.rs

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

Structure Vector
    Public X As Integer
    Public Y As Integer
    
    Public Sub New(x As Integer, y As Integer)
        Me.X = x
        Me.Y = y
    End Sub
    
    Public Shared Operator +(v1 As Vector, v2 As Vector) As Vector
        Return New Vector(v1.X + v2.X, v1.Y + v2.Y)
    End Operator
    
    Public Shared Operator =(v1 As Vector, v2 As Vector) As Boolean
        Return v1.X = v2.X AndAlso v1.Y = v2.Y
    End Operator
    
    Public Shared Operator <>(v1 As Vector, v2 As Vector) As Boolean
        Return Not (v1 = v2)
    End Operator
End Structure

Module M
    Sub Main()
        Dim v1 As New Vector(1, 2)
        Dim v2 As New Vector(3, 4)
        Dim v3 = v1 + v2
        
        __Check(CStr(v3.X), "4")
        __Check(CStr(v3.Y), "6")
        __Check(CStr(v3 = New Vector(4, 6)), "True")
    End Sub
End Module

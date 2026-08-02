' vybe-test: vb/vb_trycast_reference_types/test_vb_trycast_does_not_invoke_custom_conversion_operators
' origin: languages/vb/tests/vb/test_vb_trycast_reference_types.rs

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

Class ComplexNum
    Public Real As Double
    Public Sub New(r As Double)
        Real = r
    End Sub

    Public Shared Narrowing Operator CType(d As Double) As ComplexNum
        Return New ComplexNum(d)
    End Narrowing Operator
End Class

Module Program
    Sub Main()
        Dim dblObj As Object = 10.5
        ' TryCast only checks inheritance/interface, does not call CType operator!
        Dim cn As ComplexNum = TryCast(dblObj, ComplexNum)
        __Check(CStr(cn Is Nothing), "True")
    End Sub
End Module

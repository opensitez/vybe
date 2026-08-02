' vybe-test: vb/vb_isnothing_operator/isnothing_operator
' origin: languages/vb/tests/vb/test_vb_isnothing_operator.rs

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

Module M
    Sub Main()
        Dim obj As Object = Nothing
        Dim obj2 As New Object()
        
        ' Legacy IsNothing function vs Is Nothing operator
        __Check(CStr(IsNothing(obj)), "True")
        __Check(CStr(obj Is Nothing), "True")
        
        __Check(CStr(IsNothing(obj2)), "False")
        __Check(CStr(obj2 IsNot Nothing), "True")
    End Sub
End Module

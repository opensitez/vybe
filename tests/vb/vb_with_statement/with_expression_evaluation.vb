' vybe-test: vb/vb_with_statement/with_expression_evaluation
' origin: languages/vb/tests/vb/test_vb_with_statement.rs

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

Class Box
    Public Value As Integer
End Class

Module M
    Dim evalCount As Integer = 0
    
    Function GetBox() As Box
        evalCount = evalCount + 1
        Dim b As New Box()
        b.Value = 10
        Return b
    End Function

    Sub Main()
        With GetBox()
            .Value = .Value * 2
            __Check(CStr(.Value), "20")
        End With
        __Check(CStr(evalCount), "1")
    End Sub
End Module

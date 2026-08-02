' vybe-test: vb/vb_isnot_operator_null_checks/test_vb_isnot_operator_short_circuit_andalso
' origin: languages/vb/tests/vb/test_vb_isnot_operator_null_checks.rs

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

Module Program
    Class Container
        Public Property Content As String = "Data"
    End Class

    Sub Main()
        Dim c As Container = Nothing
        Dim hasData = (c IsNot Nothing) AndAlso (c.Content = "Data")
        __Check(CStr(hasData), "False")
    End Sub
End Module

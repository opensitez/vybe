' vybe-test: vb/vb_nullable_types/nullable_value_types_basic
' origin: languages/vb/tests/vb/test_vb_nullable_types.rs

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
        ' The ? modifier makes a value type nullable
        Dim x As Integer? = Nothing
        Dim y As Integer? = 10
        
        __Check(CStr(x.HasValue), "False")
        __Check(CStr(y.HasValue), "True")
        __Check(CStr(y.Value), "10")
    End Sub
End Module

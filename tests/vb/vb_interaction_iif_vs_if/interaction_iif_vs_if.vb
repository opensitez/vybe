' vybe-test: vb/vb_interaction_iif_vs_if/interaction_iif_vs_if
' origin: languages/vb/tests/vb/test_vb_interaction_iif_vs_if.rs

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
        Dim condition As Boolean = True
        
        ' IIf is a legacy function that evaluates BOTH true and false arguments
        ' (Not short-circuited!)
        Dim result1 = IIf(condition, "Yes", "No")
        __Check(CStr(result1), "Yes")
        
        ' If operator is short-circuited and type-safe
        Dim result2 = If(condition, "Yes", "No")
        __Check(CStr(result2), "Yes")
        
        ' If operator with two arguments acts like coalesce (expr1 ?? expr2)
        Dim val1 As String = Nothing
        Dim val2 As String = "Default"
        __Check(CStr(If(val1, val2)), "Default")
    End Sub
End Module

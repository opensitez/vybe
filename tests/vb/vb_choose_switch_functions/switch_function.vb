' vybe-test: vb/vb_choose_switch_functions/switch_function
' origin: languages/vb/tests/vb/test_vb_choose_switch_functions.rs

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
        Dim age As Integer = 25
        
        ' Switch evaluates a list of expressions and returns the corresponding value for the first True expression
        Dim category = Switch(
            age < 18, "Minor",
            age >= 18 AndAlso age < 65, "Adult",
            age >= 65, "Senior"
        )
        
        __Check(CStr(category), "Adult")
    End Sub
End Module

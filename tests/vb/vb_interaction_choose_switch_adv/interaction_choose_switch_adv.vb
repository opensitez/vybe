' vybe-test: vb/vb_interaction_choose_switch_adv/interaction_choose_switch_adv
' origin: languages/vb/tests/vb/test_vb_interaction_choose_switch_adv.rs

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
        Dim index As Integer = 2
        
        ' Choose evaluates index and returns the 1-based matching argument
        Dim choice = Choose(index, "A", "B", "C")
        __Check(CStr(choice), "B")
        
        ' Switch returns the value associated with the first True expression
        Dim val As Integer = 10
        Dim category = Switch(
            val < 0, "Negative",
            val = 0, "Zero",
            val > 0, "Positive"
        )
        __Check(CStr(category), "Positive")
    End Sub
End Module

' vybe-test: vb/vb_interaction_choose/interaction_choose_function
' origin: languages/vb/tests/vb/test_vb_interaction_choose.rs

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
        ' Choose is 1-based index
        Dim choice As String = CStr(Choose(2, "Apple", "Banana", "Cherry"))
        __Check(CStr(choice), "Banana")
        
        ' Out of bounds returns Nothing (null)
        Dim invalidChoice As Object = Choose(4, "Apple", "Banana", "Cherry")
        __Check(CStr(IsNothing(invalidChoice)), "True")
    End Sub
End Module

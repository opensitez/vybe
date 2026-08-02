' vybe-test: vb/vb_spec_information_functions/info_spec_iif_evaluates_both_branches_before_selection
' origin: languages/vb/tests/vb/test_vb_spec_information_functions.rs

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
    Function LeftValue() As String
        __Check(CStr("left"), "left")
        Return "yes"
    End Function
    Function RightValue() As String
        __Check(CStr("right"), "right")
        Return "no"
    End Function
    Sub Main()
        __Check(CStr(IIf(True, LeftValue(), RightValue())), "yes")
    End Sub
End Module

' vybe-test: vb/vb_static_locals/static_local_two_string_functions_do_not_share_state
' origin: languages/vb/tests/vb/test_vb_static_locals.rs

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
    Function LeftText(value As String) As String
        Static text As String = "L"
        text = text & value
        Return text
    End Function

    Function RightText(value As String) As String
        Static text As String = "R"
        text = text & value
        Return text
    End Function

    Sub Main()
        __Check(CStr(LeftText("a")), "La")
        __Check(CStr(RightText("b")), "Rb")
        __Check(CStr(LeftText("c")), "Lac")
        __Check(CStr(RightText("d")), "Rbd")
    End Sub
End Module

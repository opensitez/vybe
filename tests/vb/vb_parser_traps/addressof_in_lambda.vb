' vybe-test: vb/vb_parser_traps/addressof_in_lambda
' origin: languages/vb/tests/vb/test_vb_parser_traps.rs

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
    Sub Target()
        __Check(CStr("Target"), "Target")
    End Sub

    Sub Main()
        ' Sub() AddressOf Method is invalid as AddressOf returns a delegate.
        ' However, we can use it to assign to an explicit delegate inside.
        Dim act = Sub()
                      Dim d As Action = AddressOf Target
                      d()
                  End Sub
        act()
    End Sub
End Module

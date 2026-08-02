' vybe-test: vb/vb_operators_addressof/operator_addressof
' origin: languages/vb/tests/vb/test_vb_operators_addressof.rs

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
    Sub PrintMessage(msg As String)
        __Check(CStr("Message: " & msg), "Message: Hello World")
    End Sub

    Sub Main()
        ' AddressOf creates a delegate pointing to the specified procedure
        Dim del As Action(Of String) = AddressOf PrintMessage
        del("Hello World")
    End Sub
End Module

' vybe-test: vb/vb_addressof_lambda/addressof_lambda
' origin: languages/vb/tests/vb/test_vb_addressof_lambda.rs

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
    Sub Execute(action As Action)
        action()
    End Sub

    Sub Main()
        ' While lambdas don't strictly need AddressOf,
        ' sometimes they are used with delegates
        Dim a As Action = Sub() __Check(CStr("Action executed"), "Action executed")
        Execute(a)
        
        ' AddressOf can sometimes be used to refer to named methods and passed where a delegate is expected
        Execute(AddressOf PrintMessage)
    End Sub
    
    Sub PrintMessage()
        __Check(CStr("Message printed"), "Message printed")
    End Sub
End Module

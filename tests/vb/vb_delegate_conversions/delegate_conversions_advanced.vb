' vybe-test: vb/vb_delegate_conversions/delegate_conversions_advanced
' origin: languages/vb/tests/vb/test_vb_delegate_conversions.rs

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

Delegate Sub StringAction(s As String)

Module M
    Sub ExecuteAction(action As StringAction)
        action("Test Message")
    End Sub

    Sub Main()
        ' Relaxed delegate conversion allows ignoring parameters
        Dim action1 As StringAction = AddressOf HandleWithNoArgs
        action1("Ignore me")
        
        ' Or we can pass an anonymous method that takes no arguments
        ExecuteAction(Sub() __Check(CStr("Action executed with no args"), "HandleWithNoArgs called"))
    End Sub
    
    Sub HandleWithNoArgs()
        __Check(CStr("HandleWithNoArgs called"), "Action executed with no args")
    End Sub
End Module

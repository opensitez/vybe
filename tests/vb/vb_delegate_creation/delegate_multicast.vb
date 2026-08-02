' vybe-test: vb/vb_delegate_creation/delegate_multicast
' origin: languages/vb/tests/vb/test_vb_delegate_creation.rs

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

Delegate Sub Log(msg As String)

Module M
    Sub Log1(msg As String)
        __Check(CStr("1: " & msg), "1: Test")
    End Sub
    
    Sub Log2(msg As String)
        __Check(CStr("2: " & msg), "2: Test")
    End Sub

    Sub Main()
        Dim logger As Log = AddressOf Log1
        logger = CType([Delegate].Combine(logger, New Log(AddressOf Log2)), Log)
        logger("Test")
    End Sub
End Module

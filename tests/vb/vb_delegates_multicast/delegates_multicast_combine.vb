' vybe-test: vb/vb_delegates_multicast/delegates_multicast_combine
' origin: languages/vb/tests/vb/test_vb_delegates_multicast.rs

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

Delegate Sub LogAction(msg As String)

Module M
    Sub LogToConsole(msg As String)
        __Check(CStr("Console: " & msg), "Console: Test")
    End Sub
    
    Sub LogToFile(msg As String)
        __Check(CStr("File: " & msg), "File: Test")
    End Sub

    Sub Main()
        Dim d1 As LogAction = AddressOf LogToConsole
        Dim d2 As LogAction = AddressOf LogToFile
        
        ' Multicast delegate combination
        Dim d3 As LogAction = CType([Delegate].Combine(d1, d2), LogAction)
        
        d3("Test")
    End Sub
End Module

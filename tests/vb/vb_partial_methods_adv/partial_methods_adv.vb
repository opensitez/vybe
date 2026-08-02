' vybe-test: vb/vb_partial_methods_adv/partial_methods_adv
' origin: languages/vb/tests/vb/test_vb_partial_methods_adv.rs

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

Partial Class Processor
    ' Declaration of a partial method
    Partial Private Sub Log(msg As String)
    End Sub
    
    Public Sub Run()
        Log("Running")
        __Check(CStr("Done"), "Log: Running")
    End Sub
End Class

Partial Class Processor
    ' Implementation of the partial method
    Private Sub Log(msg As String)
        __Check(CStr("Log: " & msg), "Done")
    End Sub
End Class

Module M
    Sub Main()
        Dim p As New Processor()
        p.Run()
    End Sub
End Module

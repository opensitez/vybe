' vybe-test: vb/vb_control_flow_edges/try_catch_when_complex
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

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
    Function LogError() As Boolean
        __Check(CStr("Filtered"), "Filtered")
        Return True
    End Function

    Sub Main()
        Try
            Throw New System.Exception("Test")
        Catch ex As System.Exception When LogError()
            __Check(CStr("Caught"), "Caught")
        End Try
    End Sub
End Module

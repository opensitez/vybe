' vybe-test: vb/vb_control_flow_edges/return_from_catch
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
    Function Test() As String
        Try
            Throw New System.Exception()
        Catch
            Return "Caught"
        End Try
        Return "Missed"
    End Function

    Sub Main()
        __Check(CStr(Test()), "Caught")
    End Sub
End Module

' vybe-test: vb/vb_info_err_erl/info_err_erl
' origin: languages/vb/tests/vb/test_vb_info_err_erl.rs

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
    Sub Main()
        On Error Resume Next
        
10:     Dim a = 1
20:     Error 5 ' Simulate an error on line 20
        
        ' Err object contains information about run-time errors
        __Check(CStr(Err.Number), "5")
        
        ' Erl function returns the line number where the error occurred
        __Check(CStr(Erl()), "20")
        
        Err.Clear()
        __Check(CStr(Err.Number), "0")
    End Sub
End Module

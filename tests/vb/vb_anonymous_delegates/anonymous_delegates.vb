' vybe-test: vb/vb_anonymous_delegates/anonymous_delegates
' origin: languages/vb/tests/vb/test_vb_anonymous_delegates.rs

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
        ' Anonymous Sub delegate
        Dim log = Sub(msg As String) __Check(CStr("Log: " & msg), "Log: Test")
        
        ' Anonymous Function delegate
        Dim multiply = Function(x As Integer, y As Integer) As Integer
                           Return x * y
                       End Function
        
        log("Test")
        __Check(CStr(multiply(3, 4)), "12")
    End Sub
End Module

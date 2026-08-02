' vybe-test: vb/vb_call_statement/call_statement_basic
' origin: languages/vb/tests/vb/test_vb_call_statement.rs

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
        __Check(CStr(msg), "Hello using Call")
    End Sub

    Function GetValue() As Integer
        __Check(CStr("Side effect"), "Side effect")
        Return 42
    End Function

    Sub Main()
        ' The Call keyword allows calling a Sub or Function, discarding the return value if any
        Call PrintMessage("Hello using Call")
        Call GetValue()
    End Sub
End Module

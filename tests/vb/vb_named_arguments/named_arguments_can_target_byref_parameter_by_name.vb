' vybe-test: vb/vb_named_arguments/named_arguments_can_target_byref_parameter_by_name
' origin: languages/vb/tests/vb/test_vb_named_arguments.rs

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
    Sub SetValue(ByRef value As Integer, amount As Integer)
        value = amount
    End Sub

    Sub Main()
        Dim total As Integer = 0
        SetValue(amount:=9, value:=total)
        __Check(CStr(total), "9")
    End Sub
End Module

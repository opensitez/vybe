' vybe-test: vb/vb_control_flow_edges/with_statement_struct
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

Structure S
    Public Val As Integer
End Structure

Module M
    Sub Main()
        Dim s1 As New S()
        With s1
            .Val = 42
        End With
        ' Because it's a value type, With creates a copy or modifies original depending on context.
        ' Actually in VB, With on a variable Modifies the variable!
        __Check(CStr(s1.Val), "42")
    End Sub
End Module

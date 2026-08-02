' vybe-test: vb/vb_static_locals/static_local_can_remember_last_argument
' origin: languages/vb/tests/vb/test_vb_static_locals.rs

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
    Function Remember(value As Integer) As String
        Static previous As Integer = -1
        Dim result As String = previous & "->" & value
        previous = value
        Return result
    End Function

    Sub Main()
        __Check(CStr(Remember(4)), "-1->4")
        __Check(CStr(Remember(9)), "4->9")
        __Check(CStr(Remember(2)), "9->2")
    End Sub
End Module

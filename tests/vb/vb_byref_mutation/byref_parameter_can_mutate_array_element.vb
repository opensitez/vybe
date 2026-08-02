' vybe-test: vb/vb_byref_mutation/byref_parameter_can_mutate_array_element
' origin: languages/vb/tests/vb/test_vb_byref_mutation.rs

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
    Sub Bump(ByRef value As Integer)
        value = value + 5
    End Sub

    Sub Main()
        Dim values() As Integer = {1, 2, 3}
        Bump(values(1))
        __Check(CStr(values(1)), "7")
    End Sub
End Module

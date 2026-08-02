' vybe-test: vb/vb_byref_mutation/byref_can_rebind_array_variable_not_only_element
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
    Sub ReplaceNumbers(ByRef values() As Integer)
        values = New Integer() {4, 5, 6}
    End Sub

    Sub Main()
        Dim values() As Integer = New Integer() {1, 2, 3}
        ReplaceNumbers(values)
        __Check(CStr(values.Length), "3")
        __Check(CStr(values(0) + values(2)), "10")
    End Sub
End Module

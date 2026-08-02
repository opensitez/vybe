' vybe-test: vb/vb_byref_mutation/byref_can_increment_single_array_element_through_alias
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
    Sub Increment(ByRef item As Integer)
        item += 1
    End Sub

    Sub Main()
        Dim values() As Integer = New Integer() {1, 2, 3}
        Increment(values(1))
        __Check(CStr(values(1)), "3")
    End Sub
End Module

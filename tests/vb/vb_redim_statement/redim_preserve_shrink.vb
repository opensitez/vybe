' vybe-test: vb/vb_redim_statement/redim_preserve_shrink
' origin: languages/vb/tests/vb/test_vb_redim_statement.rs

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
        Dim arr() As Integer = {1, 2, 3, 4, 5}
        ReDim Preserve arr(2)
        __Check(CStr(arr.Length), "3")
        __Check(CStr(arr(1)), "2")
    End Sub
End Module

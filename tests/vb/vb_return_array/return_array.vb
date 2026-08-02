' vybe-test: vb/vb_return_array/return_array
' origin: languages/vb/tests/vb/test_vb_return_array.rs

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
    ' Method returning an array
    Function GetNames() As String()
        Return {"Alice", "Bob"}
    End Function

    Sub Main()
        Dim names = GetNames()
        __Check(CStr(names(0)), "Alice")
        __Check(CStr(names(1)), "Bob")
    End Sub
End Module

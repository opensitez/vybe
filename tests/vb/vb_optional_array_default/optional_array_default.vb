' vybe-test: vb/vb_optional_array_default/optional_array_default
' origin: languages/vb/tests/vb/test_vb_optional_array_default.rs

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
    ' Arrays cannot be Optional with a default value other than Nothing.
    ' We will test parsing of Nothing as default.
    Sub PrintFirst(Optional arr() As Integer = Nothing)
        If arr IsNot Nothing Then
            __Check(CStr(arr(0)), "Empty")
        Else
            __Check(CStr("Empty"), "5")
        End If
    End Sub

    Sub Main()
        PrintFirst()
        PrintFirst({5})
    End Sub
End Module

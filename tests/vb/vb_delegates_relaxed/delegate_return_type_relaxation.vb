' vybe-test: vb/vb_delegates_relaxed/delegate_return_type_relaxation
' origin: languages/vb/tests/vb/test_vb_delegates_relaxed.rs

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
    Delegate Function Provider() As Object

    Function ProvideString() As String
        Return "Hello"
    End Function

    Sub Main()
        ' String narrows/widens to Object, so this is valid relaxed binding
        Dim p As Provider = AddressOf ProvideString
        __Check(CStr(p().ToString()), "Hello")
    End Sub
End Module

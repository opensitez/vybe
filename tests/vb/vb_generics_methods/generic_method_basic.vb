' vybe-test: vb/vb_generics_methods/generic_method_basic
' origin: languages/vb/tests/vb/test_vb_generics_methods.rs

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
    Function Identity(Of T)(value As T) As T
        Return value
    End Function

    Sub Main()
        __Check(CStr(Identity(Of String)("Test")), "Test")
        __Check(CStr(Identity(Of Integer)(123)), "123")
    End Sub
End Module

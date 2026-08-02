' vybe-test: vb/vb_exception_nested/exception_nested_try_catch
' origin: languages/vb/tests/vb/test_vb_exception_nested.rs

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
        Try
            __Check(CStr("Outer Try"), "Outer Try")
            Try
                Throw New Exception("Inner")
            Catch ex As Exception
                __Check(CStr("Caught Inner"), "Caught Inner")
                Throw New Exception("Outer")
            Finally
                __Check(CStr("Inner Finally"), "Inner Finally")
            End Try
        Catch ex As Exception
            __Check(CStr("Caught Outer: " & ex.Message), "Caught Outer: Outer")
        Finally
            __Check(CStr("Outer Finally"), "Outer Finally")
        End Try
    End Sub
End Module

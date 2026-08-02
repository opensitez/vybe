' vybe-test: vb/vb_named_tuple_nested_structures/test_vb_named_tuple_optional_method_parameters
' origin: languages/vb/tests/vb/test_vb_named_tuple_nested_structures.rs

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

Module Program
    Private Function GetConfig(Optional cfg As (Host As String, Port As Integer) = Nothing) As String
        If cfg.Host Is Nothing Then Return "default:8080"
        Return cfg.Host & ":" & cfg.Port
    End Function

    Sub Main()
        __Check(CStr(GetConfig() & "|" & GetConfig(("localhost", 9000))), "default:8080|localhost:9000")
    End Sub
End Module

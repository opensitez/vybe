' vybe-test: vb/vb_spec_error_handling_resources/error_spec_using_can_wrap_existing_expression_result
' origin: languages/vb/tests/vb/test_vb_spec_error_handling_resources.rs

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

Class Probe
    Implements IDisposable
    Public Shared Function Build() As Probe
        Return New Probe()
    End Function
    Public Sub Dispose() Implements IDisposable.Dispose
        __Check(CStr("disposed"), "body")
    End Sub
End Class
Module M
    Sub Main()
        Using value As Probe = Probe.Build()
            __Check(CStr("body"), "disposed")
        End Using
    End Sub
End Module

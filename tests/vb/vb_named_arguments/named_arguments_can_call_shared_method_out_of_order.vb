' vybe-test: vb/vb_named_arguments/named_arguments_can_call_shared_method_out_of_order
' origin: languages/vb/tests/vb/test_vb_named_arguments.rs

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

Class Formatter
    Public Shared Function Wrap(value As String, prefix As String, suffix As String) As String
        Return prefix & value & suffix
    End Function
End Class

Module M
    Sub Main()
        __Check(CStr(Formatter.Wrap(suffix:="]", value:="core", prefix:="[")), "[core]")
    End Sub
End Module

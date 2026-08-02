' vybe-test: vb/vb_named_arguments/named_arguments_can_call_instance_method_out_of_order
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
    Public Function JoinParts(left As String, middle As String, right As String) As String
        Return left & "-" & middle & "-" & right
    End Function
End Class

Module M
    Sub Main()
        Dim formatter As New Formatter()
        __Check(CStr(formatter.JoinParts(right:="finish", left:="start", middle:="middle")), "start-middle-finish")
    End Sub
End Module

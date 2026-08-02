' vybe-test: vb/vb_generic_delegate_type_args/test_vb_generic_delegate_in_generic_class
' origin: languages/vb/tests/vb/test_vb_generic_delegate_type_args.rs

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

Class EventSource(Of T)
    Public Delegate Sub CustomHandler(sender As Object, data As T)
    Public Event OnData As CustomHandler
    Public Sub Fire(d As T)
        RaiseEvent OnData(Me, d)
    End Sub
End Class

Module Program
    Sub Main()
        Dim src As New EventSource(Of String)()
        AddHandler src.OnData, Sub(s, data) __Check(CStr("Event: " & data), "Event: Payload")
        src.Fire("Payload")
    End Sub
End Module

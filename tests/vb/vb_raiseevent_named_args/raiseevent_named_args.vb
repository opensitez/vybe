' vybe-test: vb/vb_raiseevent_named_args/raiseevent_named_args
' origin: languages/vb/tests/vb/test_vb_raiseevent_named_args.rs

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

Class Publisher
    Public Event Notify(msg As String, code As Integer)
    
    Public Sub DoNotify()
        ' RaiseEvent with named arguments
        RaiseEvent Notify(code:=100, msg:="Alert")
    End Sub
End Class

Module M
    Sub Main()
        Dim p As New Publisher()
        AddHandler p.Notify, Sub(m, c) __Check(CStr(m & c), "Alert100")
        p.DoNotify()
    End Sub
End Module

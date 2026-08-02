' vybe-test: vb/vb_events/addhandler_removehandler
' origin: languages/vb/tests/vb/test_vb_events.rs

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

Class Button
    Public Event Click()
    Public Sub DoClick()
        RaiseEvent Click()
    End Sub
End Class

Module M
    Sub OnClick()
        __Check(CStr("clicked"), "clicked")
    End Sub
    Sub Main()
        Dim btn As New Button()
        AddHandler btn.Click, AddressOf OnClick
        btn.DoClick()
        RemoveHandler btn.Click, AddressOf OnClick
        btn.DoClick()
        __Check(CStr("done"), "done")
    End Sub
End Module

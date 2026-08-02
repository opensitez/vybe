' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_withevents_property_setter_access_level
' origin: languages/vb/tests/vb/test_vb_event_subscribing_in_constructor.rs

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

Imports System

Class Publisher
    Public Event Notice As EventHandler
    Public Sub Fire()
        RaiseEvent Notice(Me, EventArgs.Empty)
    End Sub
End Class

Class EncapsulatedListener
    Public WithEvents Pub As Publisher

    Private Sub OnNotice(sender As Object, e As EventArgs) Handles Pub.Notice
        __Check(CStr("Notice Handled"), "Notice Handled")
    End Sub
End Class

Module Program
    Sub Main()
        Dim l As New EncapsulatedListener()
        Dim p As New Publisher()
        l.Pub = p
        p.Fire()
    End Sub
End Module

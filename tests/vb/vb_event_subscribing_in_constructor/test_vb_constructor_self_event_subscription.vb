' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_constructor_self_event_subscription
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

Class SelfNotifyingWidget
    Public Event InternalStateChanged As EventHandler

    Public Sub New()
        AddHandler InternalStateChanged, Sub(s, e) __Check(CStr("Self Notification Received"), "Self Notification Received")
    End Sub

    Public Sub Mutate()
        RaiseEvent InternalStateChanged(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim w As New SelfNotifyingWidget()
        w.Mutate()
    End Sub
End Module

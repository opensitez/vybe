' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_constructor_subscription_lambda_closure
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

Class LambdaSubscriber
    Public Property SignalCount As Integer = 0
    Public Sub New(pub As Source)
        AddHandler pub.Ping, Sub(s, e) SignalCount += 1
    End Sub
End Class

Class Source
    Public Event Ping As EventHandler
    Public Sub Fire()
        RaiseEvent Ping(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As New Source()
        Dim ls As New LambdaSubscriber(s)
        s.Fire()
        s.Fire()
        __Check(CStr("Signals Received: " & ls.SignalCount), "Signals Received: 2")
    End Sub
End Module

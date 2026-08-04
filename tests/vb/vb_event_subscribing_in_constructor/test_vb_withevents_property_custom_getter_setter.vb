' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_withevents_property_custom_getter_setter
' origin: languages/vb/tests/vb/test_vb_event_subscribing_in_constructor.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.
'
' Output is COLLECTED, not paired. The emitter rewrites every
' `Console.WriteLine(x)` into `__P(CStr(x))` and compares the whole output once
' at the end of `Sub Main`. Pairing the i-th print with the i-th expected line
' cannot assert anything about a loop, and loops alone were 402 of VB's 6,671
' cases.
'
' Rendering happens at the CALL SITE via `CStr`, where the expression still has
' its static type — the same reason the C# harness renders with `.ToString()`
' rather than inside the helper.

Module VybeCheck
    Public __buf As String = ""

    Sub __P(s As String)
        __buf = __buf & s & vbLf
    End Sub

    Sub __Pr(s As String)
        __buf = __buf & s
    End Sub

    ' The final WriteLine contributes a trailing newline that the expected line
    ' vector never carried, so BOTH forms are accepted.
    Sub __Check(want As String)
        If __buf <> want AndAlso __buf <> want & vbLf Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & __buf & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Imports System

Class Notifier
    Public Event Ping As EventHandler
    Public Sub Fire()
        RaiseEvent Ping(Me, EventArgs.Empty)
    End Sub
End Class

Class ExplicitPropertyListener
    Private _notifier As Notifier

    Public Custom WithEvents Property NotifierProp As Notifier
        Get
            Return _notifier
        End Get
        Set(value As Notifier)
            If _notifier IsNot Nothing Then
                RemoveHandler _notifier.Ping, AddressOf OnPing
            End If
            _notifier = value
            If _notifier IsNot Nothing Then
                AddHandler _notifier.Ping, AddressOf OnPing
            End If
        End Set
    End Property

    Private Sub OnPing(sender As Object, e As EventArgs)
        __P(CStr("Explicit Property Handled Ping"))
    End Sub
End Class

Module Program
    Sub Main()
        Dim n As New Notifier()
        Dim listener As New ExplicitPropertyListener()
        listener.NotifierProp = n
        n.Fire()
        __Check("Explicit Property Handled Ping")
    End Sub
End Module

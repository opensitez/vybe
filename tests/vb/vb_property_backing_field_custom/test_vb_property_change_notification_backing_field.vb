' vybe-test: vb/vb_property_backing_field_custom/test_vb_property_change_notification_backing_field
' origin: languages/vb/tests/vb/test_vb_property_backing_field_custom.rs

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

Class NotifyingItem
    Private _val As Integer
    Public Event ValueChanged(oldV As Integer, newV As Integer)

    Public Property Value As Integer
        Get
            Return _val
        End Get
        Set(val As Integer)
            If _val <> val Then
                Dim old As Integer = _val
                _val = val
                RaiseEvent ValueChanged(old, val)
            End If
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim item As New NotifyingItem()
        AddHandler item.ValueChanged, Sub(oldV, newV)
            __P(CStr("Changed: " & oldV & "->" & newV))
            __Check("Changed: 0->10
Changed: 10->20")
        End Sub
        item.Value = 10
        item.Value = 10 ' No event
        item.Value = 20
    End Sub
End Module

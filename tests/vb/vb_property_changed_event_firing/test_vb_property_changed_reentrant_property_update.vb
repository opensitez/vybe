' vybe-test: vb/vb_property_changed_event_firing/test_vb_property_changed_reentrant_property_update
' origin: languages/vb/tests/vb/test_vb_property_changed_event_firing.rs

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

Imports System.ComponentModel
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


Class ReentrantVM
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _val1 As Integer
    Private _val2 As Integer

    Public Property Val1 As Integer
        Get
            Return _val1
        End Get
        Set(v As Integer)
            _val1 = v
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Val1"))
        End Set
    End Property

    Public Property Val2 As Integer
        Get
            Return _val2
        End Get
        Set(v As Integer)
            _val2 = v
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Val2"))
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim vm As New ReentrantVM()
        AddHandler vm.PropertyChanged, Sub(s, e)
            If e.PropertyName = "Val1" Then
                vm.Val2 = vm.Val1 * 10
            End If
            __Check("50")
        End Sub

        vm.Val1 = 5
        __P(CStr(vm.Val2))
    End Sub
End Module

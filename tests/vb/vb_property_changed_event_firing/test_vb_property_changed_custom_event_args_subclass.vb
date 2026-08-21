' vybe-test: vb/vb_property_changed_event_firing/test_vb_property_changed_custom_event_args_subclass
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


Class ExtendedPropertyChangedEventArgs
    Inherits PropertyChangedEventArgs
    Public Property OldValue As Object
    Public Property NewValue As Object
    Public Sub New(propName As String, oldVal As Object, newVal As Object)
        MyBase.New(propName)
        OldValue = oldVal
        NewValue = newVal
    End Sub
End Class

Class RichModel
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _status As String = "Init"
    Public Property Status As String
        Get
            Return _status
        End Get
        Set(value As String)
            Dim old = _status
            _status = value
            RaiseEvent PropertyChanged(Me, New ExtendedPropertyChangedEventArgs("Status", old, value))
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim rm As New RichModel()
        AddHandler rm.PropertyChanged, Sub(s, e)
            Dim ext = CType(e, ExtendedPropertyChangedEventArgs)
            __P(CStr(ext.PropertyName & ": " & ext.OldValue.ToString() & " -> " & ext.NewValue.ToString()))
            __Check("Status: Init -> Ready")
        End Sub
        rm.Status = "Ready"
    End Sub
End Module

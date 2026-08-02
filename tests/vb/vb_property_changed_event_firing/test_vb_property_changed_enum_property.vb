' vybe-test: vb/vb_property_changed_event_firing/test_vb_property_changed_enum_property
' origin: languages/vb/tests/vb/test_vb_property_changed_event_firing.rs

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

Imports System.ComponentModel

Enum NetworkState
    Disconnected
    Connected
End Enum

Class Device
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _state As NetworkState
    Public Property State As NetworkState
        Get
            Return _state
        End Get
        Set(value As NetworkState)
            _state = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("State"))
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim dev As New Device()
        AddHandler dev.PropertyChanged, Sub(s, e) __Check(CStr(e.PropertyName & "=" & dev.State.ToString()), "State=Connected")
        dev.State = NetworkState.Connected
    End Sub
End Module

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

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Imports System.ComponentModel

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
            __Check(CStr(ext.PropertyName & ": " & ext.OldValue.ToString() & " -> " & ext.NewValue.ToString()), "Status: Init -> Ready")
        End Sub
        rm.Status = "Ready"
    End Sub
End Module

' vybe-test: vb/vb_property_changed_event_firing/test_vb_inotify_property_changing_interface
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

Class EditableItem
    Implements INotifyPropertyChanging, INotifyPropertyChanged
    Public Event PropertyChanging As PropertyChangingEventHandler Implements INotifyPropertyChanging.PropertyChanging
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _score As Integer
    Public Property Score As Integer
        Get
            Return _score
        End Get
        Set(value As Integer)
            If _score <> value Then
                RaiseEvent PropertyChanging(Me, New PropertyChangingEventArgs("Score"))
                _score = value
                RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Score"))
            End If
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim item As New EditableItem()
        AddHandler item.PropertyChanging, Sub(s, e) __Check(CStr("Changing:" & e.PropertyName), "Changing:Score")
        AddHandler item.PropertyChanged, Sub(s, e) __Check(CStr("Changed:" & e.PropertyName), "Changed:Score")
        item.Score = 95
    End Sub
End Module

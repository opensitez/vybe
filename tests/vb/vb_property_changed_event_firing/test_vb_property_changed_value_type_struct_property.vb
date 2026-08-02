' vybe-test: vb/vb_property_changed_event_firing/test_vb_property_changed_value_type_struct_property
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

Structure Point2D
    Public X, Y As Integer
End Structure

Class NodeViewModel
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _pos As Point2D
    Public Property Position As Point2D
        Get
            Return _pos
        End Get
        Set(value As Point2D)
            _pos = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Position"))
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim nvm As New NodeViewModel()
        AddHandler nvm.PropertyChanged, Sub(s, e) __Check(CStr("Node Moved: " & e.PropertyName), "Node Moved: Position")
        nvm.Position = New Point2D With {.X = 10, .Y = 20}
    End Sub
End Module

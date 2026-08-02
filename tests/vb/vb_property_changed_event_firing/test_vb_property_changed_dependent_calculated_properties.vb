' vybe-test: vb/vb_property_changed_event_firing/test_vb_property_changed_dependent_calculated_properties
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

Class Employee
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _first As String
    Private _last As String

    Public Property FirstName As String
        Get
            Return _first
        End Get
        Set(value As String)
            _first = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("FirstName"))
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("FullName"))
        End Set
    End Property

    Public ReadOnly Property FullName As String
        Get
            Return _first & " " & _last
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim emp As New Employee()
        Dim firedList As New System.Collections.Generic.List(Of String)()
        AddHandler emp.PropertyChanged, Sub(s, e) firedList.Add(e.PropertyName)
        emp.FirstName = "John"
        __Check(CStr(String.Join(",", firedList)), "FirstName,FullName")
    End Sub
End Module

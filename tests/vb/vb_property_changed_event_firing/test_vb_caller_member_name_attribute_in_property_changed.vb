' vybe-test: vb/vb_property_changed_event_firing/test_vb_caller_member_name_attribute_in_property_changed
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
Imports System.Runtime.CompilerServices

Class ViewModelBase
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Protected Sub OnPropertyChanged(<CallerMemberName> Optional propName As String = Nothing)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propName))
    End Sub
End Class

Class UserViewModel
    Inherits ViewModelBase

    Private _title As String
    Public Property Title As String
        Get
            Return _title
        End Get
        Set(value As String)
            If _title <> value Then
                _title = value
                OnPropertyChanged() ' Auto infers "Title"!
            End If
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim vm As New UserViewModel()
        Dim notifiedProp = ""
        AddHandler vm.PropertyChanged, Sub(s, e) notifiedProp = e.PropertyName
        vm.Title = "Manager"
        __Check(CStr(notifiedProp), "Title")
    End Sub
End Module

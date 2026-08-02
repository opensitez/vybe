' vybe-test: vb/vb_property_changed_event_firing/test_vb_inotify_property_changed_no_event_if_same_value
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

Class Account
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _balance As Decimal = 100.0D
    Public Property Balance As Decimal
        Get
            Return _balance
        End Get
        Set(value As Decimal)
            If _balance <> value Then
                _balance = value
                RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Balance"))
            End If
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim acc As New Account()
        Dim fired = False
        AddHandler acc.PropertyChanged, Sub(s, e) fired = True
        acc.Balance = 100.0D ' Same value as initial
        __Check(CStr(fired), "False")
    End Sub
End Module

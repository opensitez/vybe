' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_custom_event_state_tracking
' origin: languages/vb/tests/vb/test_vb_custom_event_remove_nonexistent.rs

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

Imports System

Class FilteredPublisher
    Private count As Integer = 0
    Public Custom Event ItemAdded As Action
        AddHandler(value As Action)
            count += 1
        End AddHandler
        RemoveHandler(value As Action)
            count -= 1
        End RemoveHandler
        RaiseEvent()
        End RaiseEvent
    End Event
    Public Function GetCount() As Integer
        Return count
    End Function
End Class

Module Program
    Private Sub Dummy() : End Sub

    Sub Main()
        Dim fp As New FilteredPublisher()
        AddHandler fp.ItemAdded, AddressOf Dummy
        AddHandler fp.ItemAdded, AddressOf Dummy
        __Check(CStr(fp.GetCount()), "2")
        RemoveHandler fp.ItemAdded, AddressOf Dummy
        __Check(CStr(fp.GetCount()), "1")
    End Sub
End Module

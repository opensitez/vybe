' vybe-test: vb/vb_delegates_addressof/event_withevents_handles
' origin: languages/vb/tests/vb/test_vb_delegates_addressof.rs

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

Class C
Public Event E()
Public Sub Raise()
RaiseEvent E()
End Sub
End Class
Class Wrapper
Private WithEvents _c As New C()
Private Sub Handler() Handles _c.E
__Check(CStr("Handled"), "Handled")
End Sub
Public Sub Test()
_c.Raise()
End Sub
End Class
Module M
Sub Main()
Dim w As New Wrapper()
w.Test()
End Sub
End Module

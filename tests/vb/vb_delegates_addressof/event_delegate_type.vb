' vybe-test: vb/vb_delegates_addressof/event_delegate_type
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

Delegate Sub D(v As Integer)
Class C
Public Event E As D
Public Sub Raise()
RaiseEvent E(5)
End Sub
End Class
Module M
Sub Handler(v As Integer)
__Check(CStr(v), "5")
End Sub
Sub Main()
Dim c1 As New C()
AddHandler c1.E, AddressOf Handler
c1.Raise()
End Sub
End Module

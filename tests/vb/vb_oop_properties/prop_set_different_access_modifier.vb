' vybe-test: vb/vb_oop_properties/prop_set_different_access_modifier
' origin: languages/vb/tests/vb/test_vb_oop_properties.rs

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
Private _v As Integer
Public Property V As Integer
Get
Return _v
End Get
Private Set(value As Integer) ' Mixed access
_v = value
End Set
End Property
Public Sub SetV(val As Integer)
V = val
End Sub
End Class
Module M
Sub Main()
Dim c1 As New C()
c1.SetV(8)
__Check(CStr(c1.V), "8")
End Sub
End Module

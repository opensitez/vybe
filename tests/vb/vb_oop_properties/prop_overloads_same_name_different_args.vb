' vybe-test: vb/vb_oop_properties/prop_overloads_same_name_different_args
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
Public Property V(s As String) As Integer
Get
Return 1
End Get
Set(value As Integer)
End Set
End Property
Public Property V(i As Integer) As Integer
Get
Return 2
End Get
Set(value As Integer)
End Set
End Property
End Class
Module M
Sub Main()
Dim c1 As New C()
__Check(CStr(c1.V("A") + c1.V(1)), "3")
End Sub
End Module

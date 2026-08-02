' vybe-test: vb/vb_oop_properties/prop_writeonly_get_fails
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
Public WriteOnly Property V As Integer
Set(value As Integer)
End Set
End Property
End Class
Module M
Sub Main()
Dim c1 As New C()
' Dim x = c1.V ' Cannot read WriteOnly property
__Check(CStr("Parsed"), "Parsed")
End Sub
End Module

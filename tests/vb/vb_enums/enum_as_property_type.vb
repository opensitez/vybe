' vybe-test: vb/vb_enums/enum_as_property_type
' origin: languages/vb/tests/vb/test_vb_enums.rs

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

Enum E
A = 5
End Enum
Class C
Public Property V As E
End Class
Module M
Sub Main()
Dim c1 As New C()
c1.V = E.A
__Check(CStr(CInt(c1.V)), "5")
End Sub
End Module

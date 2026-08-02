' vybe-test: vb/vb_enums/enum_as_method_return_type
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
A = 10
End Enum
Module M
Function GetE() As E
Return E.A
End Function
Sub Main()
__Check(CStr(CInt(GetE())), "10")
End Sub
End Module

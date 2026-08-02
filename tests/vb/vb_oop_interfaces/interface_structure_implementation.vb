' vybe-test: vb/vb_oop_interfaces/interface_structure_implementation
' origin: languages/vb/tests/vb/test_vb_oop_interfaces.rs

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

Interface I
Function GetV() As Integer
End Interface
Structure S
Implements I
Public Function GetV() As Integer Implements I.GetV
Return 42
End Function
End Structure
Module M
Sub Main()
Dim s1 As I = New S()
__Check(CStr(s1.GetV()), "42")
End Sub
End Module

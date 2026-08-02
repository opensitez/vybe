' vybe-test: vb/vb_ctype_custom_operator/test_vb_ctype_custom_narrowing_operator
' origin: languages/vb/tests/vb/test_vb_ctype_custom_operator.rs

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

Class Temperature
    Public Celsius As Double
    Public Sub New(c As Double)
        Celsius = c
    End Sub

    ' Narrowing operator: Temperature to Integer (may lose decimal precision)
    Public Shared Narrowing Operator CType(t As Temperature) As Integer
        Return CInt(t.Celsius)
    End Shared Narrowing Operator
End Class

Module Program
    Sub Main()
        Dim temp As New Temperature(36.6)
        Dim cInt As Integer = CType(temp, Integer)
        __Check(CStr(cInt), "37")
    End Sub
End Module

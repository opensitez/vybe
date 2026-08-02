' vybe-test: vb/vb_casting_patterns/typeof_detects_value_type
' origin: languages/vb/tests/vb/test_vb_casting_patterns.rs

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

Module M
    Sub Main()
        Dim boxed As Object = 99
        __Check(CStr(TypeOf boxed Is Integer), "True")
        __Check(CStr(TypeOf boxed Is Decimal), "False")
    End Sub
End Module

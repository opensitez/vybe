' vybe-test: vb/vb_casting_patterns/typeof_identifies_reference_type
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

Class Animal
End Class

Module M
    Sub Main()
        Dim o As Object = New Animal()
        __Check(CStr(TypeOf o Is Animal), "True")
        __Check(CStr(TypeOf o IsNot Nothing), "True")
    End Sub
End Module

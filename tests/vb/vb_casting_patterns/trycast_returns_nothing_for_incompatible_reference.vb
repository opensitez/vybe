' vybe-test: vb/vb_casting_patterns/trycast_returns_nothing_for_incompatible_reference
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

Class Base
End Class

Class Other
End Class

Module M
    Sub Main()
        Dim b As Base = New Base()
        Dim d As Other = TryCast(b, Other)
        __Check(CStr(d Is Nothing), "True")
    End Sub
End Module

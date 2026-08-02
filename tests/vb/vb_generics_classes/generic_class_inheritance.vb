' vybe-test: vb/vb_generics_classes/generic_class_inheritance
' origin: languages/vb/tests/vb/test_vb_generics_classes.rs

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

Class Base(Of T)
    Public Value As T
End Class

Class DerivedInt
    Inherits Base(Of Integer)
End Class

Module M
    Sub Main()
        Dim d As New DerivedInt()
        d.Value = 99
        __Check(CStr(d.Value), "99")
    End Sub
End Module

' vybe-test: vb/vb_null_reference_exception_guards/test_vb_chained_null_conditional_calls
' origin: languages/vb/tests/vb/test_vb_null_reference_exception_guards.rs

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

Class Company
    Public Property Owner As Person
End Class

Class Person
    Public Property Address As Address
End Class

Class Address
    Public Property ZipCode As String = "90210"
End Class

Module Program
    Sub Main()
        Dim comp As Company = Nothing
        __Check(CStr(comp?.Owner?.Address?.ZipCode Is Nothing), "True")
        comp = New Company() With {.Owner = New Person() With {.Address = New Address()}}
        __Check(CStr(comp?.Owner?.Address?.ZipCode), "90210")
    End Sub
End Module

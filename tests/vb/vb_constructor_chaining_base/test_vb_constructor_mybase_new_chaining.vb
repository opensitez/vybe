' vybe-test: vb/vb_constructor_chaining_base/test_vb_constructor_mybase_new_chaining
' origin: languages/vb/tests/vb/test_vb_constructor_chaining_base.rs

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

Class BaseResource
    Public Shared InitLog As String = ""
    Public Sub New(msg As String)
        InitLog &= "Base:" & msg & ";"
    End Sub
End Class

Class DerivedResource
    Inherits BaseResource
    Public Sub New(msg As String)
        MyBase.New(msg)
        InitLog &= "Derived:" & msg & ";"
    End Sub
End Class

Module Program
    Sub Main()
        Dim r As New DerivedResource("Test")
        __Check(CStr(BaseResource.InitLog), "Base:Test;Derived:Test;")
    End Sub
End Module

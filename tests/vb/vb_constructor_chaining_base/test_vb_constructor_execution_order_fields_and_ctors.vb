' vybe-test: vb/vb_constructor_chaining_base/test_vb_constructor_execution_order_fields_and_ctors
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

Class BaseObj
    Public Field1 As String = "BaseField"
    Public Sub New()
        __Check(CStr("BaseCtor"), "BaseCtor")
    End Sub
End Class

Class DerivedObj
    Inherits BaseObj
    Public Field2 As String = "DerivedField"
    Public Sub New()
        MyBase.New()
        __Check(CStr("DerivedCtor"), "DerivedCtor")
    End Sub
End Class

Module Program
    Sub Main()
        Dim d As New DerivedObj()
        __Check(CStr(d.Field1 & ":" & d.Field2), "BaseField:DerivedField")
    End Sub
End Module

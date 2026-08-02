' vybe-test: vb/vb_reflection_constructors_create_instance/test_vb_reflection_constructor_info_invoke
' origin: languages/vb/tests/vb/test_vb_reflection_constructors_create_instance.rs

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
Imports System.Reflection

Class Product
    Public SKU As String
    Public Sub New(s As String) : SKU = s : End Sub
End Class

Module Program
    Sub Main()
        Dim t = GetType(Product)
        Dim ctor = t.GetConstructor({GetType(String)})
        Dim p As Product = CType(ctor.Invoke({"SKU-100"}), Product)
        __Check(CStr(p.SKU), "SKU-100")
    End Sub
End Module

' vybe-test: vb/vb_reflection_method_info_generic_invoke/test_vb_reflection_method_info_custom_attribute
' origin: languages/vb/tests/vb/test_vb_reflection_method_info_generic_invoke.rs

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

<AttributeUsage(AttributeTargets.Method)>
Class RouteAttribute
    Inherits Attribute
    Public Path As String
    Public Sub New(p As String) : Path = p : End Sub
End Class

Class Controller
    <Route("/api/users")>
    Public Sub GetUsers() : End Sub
End Class

Module Program
    Sub Main()
        Dim m = GetType(Controller).GetMethod("GetUsers")
        Dim attr = CType(m.GetCustomAttributes(GetType(RouteAttribute), False)(0), RouteAttribute)
        __Check(CStr(attr.Path), "/api/users")
    End Sub
End Module

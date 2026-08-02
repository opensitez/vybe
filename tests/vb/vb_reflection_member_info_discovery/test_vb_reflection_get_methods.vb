' vybe-test: vb/vb_reflection_member_info_discovery/test_vb_reflection_get_methods
' origin: languages/vb/tests/vb/test_vb_reflection_member_info_discovery.rs

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

Imports System.Reflection

Class TargetClass
    Public Sub Foo()
    End Sub
    Public Function Bar() As Integer
        Return 0
    End Function
End Class

Module Program
    Sub Main()
        Dim t As Type = GetType(TargetClass)
        Dim methods As MethodInfo() = t.GetMethods(BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.DeclaredOnly)
        __Check(CStr(methods.Length), "2")
    End Sub
End Module

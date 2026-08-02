' vybe-test: vb/vb_global_namespace_targeting/global_namespace_targeting
' origin: languages/vb/tests/vb/test_vb_global_namespace_targeting.rs

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

Namespace Root
    Public Class A
        Public Sub Show()
            __Check(CStr("Root.A"), "Nested.A")
        End Sub
    End Class
End Namespace

Namespace Nested
    Public Class A
        Public Sub Show()
            __Check(CStr("Nested.A"), "Root.A")
        End Sub
    End Class

    Module M
        Sub Main()
            Dim obj1 As New A()
            obj1.Show()
            
            Dim obj2 As New Global.Root.A()
            obj2.Show()
        End Sub
    End Module
End Namespace

' vybe-test: vb/vb_system_reflection_matrix/reflection_interfaces_discovery
' origin: languages/vb/tests/vb/test_vb_system_reflection_matrix.rs

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

Module M
    Sub Main()
        Dim impl As New Target()
        Dim types() As Type = impl.GetType().GetInterfaces()
        __Check(CStr(types.Length), "1")
        __Check(CStr(types(0).Name), "ITraceable")
    End Sub

    Interface ITraceable
        Sub Mark()
    End Interface

    Class Target
        Implements ITraceable
        Public Sub Mark() Implements ITraceable.Mark
        End Sub
    End Class
End Module

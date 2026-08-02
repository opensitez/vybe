' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_type_parameter_name_shadowing
' origin: languages/vb/tests/vb/test_vb_generic_interface_inheritance.rs

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

Interface IOuter(Of T)
    Interface IInner(Of T)
        Function Combine(o As T) As String
    End Interface
End Interface

Class Impl
    Implements IOuter(Of String).IInner(Of Integer)
    Public Function Combine(o As Integer) As String Implements IOuter(Of String).IInner(Of Integer).Combine
        Return "IntegerVal_" & o
    End Function
End Class

Module Program
    Sub Main()
        Dim impl As IOuter(Of String).IInner(Of Integer) = New Impl()
        __Check(CStr(impl.Combine(77)), "IntegerVal_77")
    End Sub
End Module

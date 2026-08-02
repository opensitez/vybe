' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_inheritance_with_additional_methods
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

Interface IService(Of T)
    Sub Execute(item As T)
End Interface

Interface IAdvancedService(Of T)
    Inherits IService(Of T)
    Sub ExecuteBatch(items As T())
End Interface

Class StringService
    Implements IAdvancedService(Of String)
    Public Sub Execute(item As String) Implements IService(Of String).Execute
        __Check(CStr("Single: " & item), "Single: One")
    End Sub
    Public Sub ExecuteBatch(items As String()) Implements IAdvancedService(Of String).ExecuteBatch
        __Check(CStr("Batch: " & String.Join("-", items)), "Batch: Two-Three")
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As IAdvancedService(Of String) = New StringService()
        s.Execute("One")
        s.ExecuteBatch({"Two", "Three"})
    End Sub
End Module

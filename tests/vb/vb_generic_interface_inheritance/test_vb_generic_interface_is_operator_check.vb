' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_is_operator_check
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

Interface ICheckable(Of T)
End Interface

Class Impl
    Implements ICheckable(Of String)
End Class

Module Program
    Sub Main()
        Dim obj As Object = New Impl()
        __Check(CStr(TypeOf obj Is ICheckable(Of String)), "True")
        __Check(CStr(TypeOf obj Is ICheckable(Of Integer)), "False")
    End Sub
End Module

' vybe-test: vb/vb_generic_type_casting_as_is/test_vb_generic_interface_casting
' origin: languages/vb/tests/vb/test_vb_generic_type_casting_as_is.rs

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

Interface IService : End Interface
Class ServiceImpl : Implements IService : End Class

Module Program
    Private Function AsInterface(Of TInterface As Class)(obj As Object) As TInterface
        Return TryCast(obj, TInterface)
    End Function

    Sub Main()
        Dim impl As Object = New ServiceImpl()
        Dim svc As IService = AsInterface(Of IService)(impl)
        __Check(CStr(svc IsNot Nothing), "True")
    End Sub
End Module

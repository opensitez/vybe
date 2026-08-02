' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_trycast_check
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
    Sub Serve(t As T)
End Interface

Class ServiceImpl
    Implements IService(Of Integer)
    Public Sub Serve(t As Integer) Implements IService(Of Integer).Serve
        __Check(CStr("Serving: " & t), "Serving: 100")
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As Object = New ServiceImpl()
        Dim s = TryCast(obj, IService(Of Integer))
        If s IsNot Nothing Then
            s.Serve(100)
        End If
    End Sub
End Module

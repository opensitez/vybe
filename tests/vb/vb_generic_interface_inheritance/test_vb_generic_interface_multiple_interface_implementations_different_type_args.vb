' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_multiple_interface_implementations_different_type_args
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

Interface IHandler(Of T)
    Sub Handle(item As T)
End Interface

Class DualHandler
    Implements IHandler(Of Integer), IHandler(Of String)
    Public Sub HandleInt(item As Integer) Implements IHandler(Of Integer).Handle
        __Check(CStr("Int: " & item), "Int: 10")
    End Sub
    Public Sub HandleString(item As String) Implements IHandler(Of String).Handle
        __Check(CStr("String: " & item), "String: Hello")
    End Sub
End Class

Module Program
    Sub Main()
        Dim dh As New DualHandler()
        Dim hInt As IHandler(Of Integer) = dh
        Dim hStr As IHandler(Of String) = dh
        hInt.Handle(10)
        hStr.Handle("Hello")
    End Sub
End Module

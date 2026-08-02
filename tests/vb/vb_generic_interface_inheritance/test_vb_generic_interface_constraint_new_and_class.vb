' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_constraint_new_and_class
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

Interface IFactory(Of T As {Class, New})
    Function CreateInstance() As T
End Interface

Class ProductFactory(Of T As {Class, New})
    Implements IFactory(Of T)
    Public Function CreateInstance() As T Implements IFactory(Of T).CreateInstance
        Return New T()
    End Function
End Class

Class Car
    Public Model As String = "Sedan"
End Class

Module Program
    Sub Main()
        Dim f As IFactory(Of Car) = New ProductFactory(Of Car)()
        Dim c = f.CreateInstance()
        __Check(CStr(c.Model), "Sedan")
    End Sub
End Module

' vybe-test: vb/vb_interfaces_basic/interface_basic_implementation
' origin: languages/vb/tests/vb/test_vb_interfaces_basic.rs

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

Interface IAnimal
    Sub Speak()
    Function GetName() As String
End Interface

Class Dog
    Implements IAnimal
    
    Public Sub Speak() Implements IAnimal.Speak
        __Check(CStr("Woof"), "Woof")
    End Sub
    
    Public Function GetName() As String Implements IAnimal.GetName
        Return "Buddy"
    End Function
End Class

Module M
    Sub Main()
        Dim a As IAnimal = New Dog()
        a.Speak()
        __Check(CStr(a.GetName()), "Buddy")
    End Sub
End Module

' vybe-test: vb/vb_mybase_new_call/mybase_new_call
' origin: languages/vb/tests/vb/test_vb_mybase_new_call.rs

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

Class Base
    Public Sub New(name As String)
        __Check(CStr("Base: " & name), "Base: Default")
    End Sub
End Class

Class Derived
    Inherits Base
    
    Public Sub New()
        ' Calling the base constructor explicitly using MyBase.New
        MyBase.New("Default")
        __Check(CStr("Derived"), "Derived")
    End Sub
End Class

Module M
    Sub Main()
        Dim d As New Derived()
    End Sub
End Module

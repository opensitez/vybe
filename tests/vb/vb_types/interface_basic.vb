' vybe-test: vb/vb_types/interface_basic
' origin: languages/vb/tests/vb/test_vb_types.rs

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

Interface IGreeter
    Function Greet() As String
End Interface

Class HelloGreeter
    Implements IGreeter
    Public Function Greet() As String
        Return "Hello!"
    End Function
End Class

Module M
    Sub Main()
        Dim g As New HelloGreeter()
        __Check(CStr(g.Greet()), "Hello!")
    End Sub
End Module

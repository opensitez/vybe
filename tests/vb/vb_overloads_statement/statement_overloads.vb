' vybe-test: vb/vb_overloads_statement/statement_overloads
' origin: languages/vb/tests/vb/test_vb_overloads_statement.rs

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
    Public Sub Process(x As Integer)
        __Check(CStr("Process Integer: " & x), "Process Integer: 10")
    End Sub
End Class

Class Derived
    Inherits Base
    
    ' Overloads is technically optional when the signatures are different,
    ' but it's used to explicitly define overloaded methods across inheritance bounds
    Public Overloads Sub Process(x As String)
        __Check(CStr("Process String: " & x), "Process String: Hello")
    End Sub
End Class

Module M
    Sub Main()
        Dim d As New Derived()
        d.Process(10)
        d.Process("Hello")
    End Sub
End Module

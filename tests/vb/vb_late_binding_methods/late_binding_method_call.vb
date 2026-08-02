' vybe-test: vb/vb_late_binding_methods/late_binding_method_call
' origin: languages/vb/tests/vb/test_vb_late_binding_methods.rs

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

Class Greeter
    Public Function SayHello(name As String) As String
        Return "Hello " & name
    End Function
End Class

Module M
    Sub Main()
        ' Using Object type forces late binding (if Option Strict is Off, which is default)
        Dim g As Object = New Greeter()
        
        ' Method call is resolved at runtime
        __Check(CStr(g.SayHello("VB")), "Hello VB")
    End Sub
End Module

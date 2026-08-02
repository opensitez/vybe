' vybe-test: vb/vb_defint_implicit_typing/defint_implicit_typing
' origin: languages/vb/tests/vb/test_vb_defint_implicit_typing.rs

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

' Variables starting with I through N default to Integer
DefInt I-N
' Variables starting with S default to String
DefStr S

Module M
    Sub Main()
        ' iVar starts with I, so it is an Integer implicitly
        Dim iVar = 10
        Dim nVar = 20
        Dim sVar = "Hello"
        
        __Check(CStr(iVar.GetType().Name), "Int32")
        __Check(CStr(sVar.GetType().Name), "String")
    End Sub
End Module

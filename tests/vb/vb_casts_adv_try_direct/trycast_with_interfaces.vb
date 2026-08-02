' vybe-test: vb/vb_casts_adv_try_direct/trycast_with_interfaces
' origin: languages/vb/tests/vb/test_vb_casts_adv_try_direct.rs

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

Interface ITest
End Interface

Class A
    Implements ITest
End Class

Class B
End Class

Module M
    Sub Main()
        Dim objA As Object = New A()
        Dim objB As Object = New B()
        
        Dim tA As ITest = TryCast(objA, ITest)
        __Check(CStr(tA IsNot Nothing), "True")
        
        Dim tB As ITest = TryCast(objB, ITest)
        __Check(CStr(tB IsNot Nothing), "False")
    End Sub
End Module

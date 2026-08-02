' vybe-test: vb/vb_nameof_advanced/nameof_operator_members
' origin: languages/vb/tests/vb/test_vb_nameof_advanced.rs

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

Class Data
    Public Property Value As Integer
End Class

Module M
    Sub Main()
        ' NameOf can reference members of a type without an instance
        __Check(CStr(NameOf(Data.Value)), "Value")
        
        Dim d As New Data()
        __Check(CStr(NameOf(d.Value)), "Value")
        
        ' NameOf with local variables
        Dim localVariable As String = ""
        __Check(CStr(NameOf(localVariable)), "localVariable")
    End Sub
End Module

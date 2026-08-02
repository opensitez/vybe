' vybe-test: vb/vb_generic_default_operator/test_vb_generic_default_value_primitives_and_references
' origin: languages/vb/tests/vb/test_vb_generic_default_operator.rs

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

Module Helper
    Public Function GetDefault(Of T)() As T
        Return CType(Nothing, T)
    End Function
End Module

Module Program
    Sub Main()
        Dim defaultInt As Integer = Helper.GetDefault(Of Integer)()
        Dim defaultBool As Boolean = Helper.GetDefault(Of Boolean)()
        Dim defaultStr As String = Helper.GetDefault(Of String)()
        __Check(CStr(defaultInt), "0")
        __Check(CStr(defaultBool), "False")
        __Check(CStr(defaultStr Is Nothing), "True")
    End Sub
End Module

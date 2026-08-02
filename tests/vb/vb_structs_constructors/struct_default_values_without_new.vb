' vybe-test: vb/vb_structs_constructors/struct_default_values_without_new
' origin: languages/vb/tests/vb/test_vb_structs_constructors.rs

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

Structure Status
    Public Code As Integer
    Public Message As String
End Structure

Module M
    Sub Main()
        ' A struct can be used without New, fields are zeroed/Nothing
        Dim s As Status
        __Check(CStr(s.Code), "0")
        __Check(CStr(IsNothing(s.Message)), "True")
    End Sub
End Module

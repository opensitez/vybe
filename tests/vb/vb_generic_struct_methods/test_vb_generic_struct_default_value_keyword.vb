' vybe-test: vb/vb_generic_struct_methods/test_vb_generic_struct_default_value_keyword
' origin: languages/vb/tests/vb/test_vb_generic_struct_methods.rs

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

Structure Holder(Of T)
    Public Item As T
    Public Function GetDefault() As T
        Return Nothing ' VB "Nothing" for generics evaluates to default(T)
    End Function
End Structure

Module Program
    Sub Main()
        Dim hInt As New Holder(Of Integer)()
        Dim hStr As New Holder(Of String)()
        __Check(CStr(hInt.GetDefault() & "|" & (hStr.GetDefault() Is Nothing)), "0|True")
    End Sub
End Module

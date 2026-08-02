' vybe-test: vb/vb_generic_struct_methods/test_vb_generic_struct_static_shared_field
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

Structure StaticStruct(Of T)
    Public Shared DefaultVal As T
    Public Item As T
    Public Sub New(i As T)
        Item = i
    End Sub
End Structure

Module Program
    Sub Main()
        StaticStruct(Of Integer).DefaultVal = -1
        StaticStruct(Of String).DefaultVal = "N/A"

        __Check(CStr(StaticStruct(Of Integer).DefaultVal & "|" & StaticStruct(Of String).DefaultVal), "-1|N/A")
    End Sub
End Module

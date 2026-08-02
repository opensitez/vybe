' vybe-test: vb/vb_reflection_method_info_generic_invoke/test_vb_reflection_is_generic_method_definition
' origin: languages/vb/tests/vb/test_vb_reflection_method_info_generic_invoke.rs

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

Class Utility
    Public Function Process(Of T)(item As T) As String : Return item.ToString() : End Function
    Public Function NonGeneric(item As String) As String : Return item : End Function
End Class

Module Program
    Sub Main()
        Dim mGen = GetType(Utility).GetMethod("Process")
        Dim mNonGen = GetType(Utility).GetMethod("NonGeneric")
        __Check(CStr(mGen.IsGenericMethodDefinition & "|" & mNonGen.IsGenericMethodDefinition), "True|False")
    End Sub
End Module

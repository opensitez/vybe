' vybe-test: vb/vb_reflection_attributes_code_gen/test_vb_reflection_make_generic_method_dynamically
' origin: languages/vb/tests/vb/test_vb_reflection_attributes_code_gen.rs

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

Imports System
Imports System.Reflection

Class Utility
    Public Shared Function Wrap(Of T)(val As T) As String
        Return "Wrapped:" & val.ToString()
    End Function
End Class

Module Program
    Sub Main()
        Dim method = GetType(Utility).GetMethod("Wrap")
        Dim genericMethod = method.MakeGenericMethod(GetType(Integer))
        Dim result = genericMethod.Invoke(Nothing, New Object() {99})
        __Check(CStr(result), "Wrapped:99")
    End Sub
End Module

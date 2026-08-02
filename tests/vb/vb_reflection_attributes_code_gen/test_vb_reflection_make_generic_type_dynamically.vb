' vybe-test: vb/vb_reflection_attributes_code_gen/test_vb_reflection_make_generic_type_dynamically
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
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim openType = GetType(List(Of ))
        Dim closedType = openType.MakeGenericType(GetType(Integer))
        Dim listInstance = Activator.CreateInstance(closedType)
        __Check(CStr(listInstance.GetType().GenericTypeArguments(0).Name), "Int32")
    End Sub
End Module

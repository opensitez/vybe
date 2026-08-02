' vybe-test: vb/vb_generic_struct_methods/test_vb_generic_struct_generic_method_inside_generic_struct
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

Structure ConverterStruct(Of TInput)
    Public InputData As TInput
    Public Sub New(input As TInput)
        InputData = input
    End Sub
    Public Function ConvertTo(Of TOutput)(converter As System.Func(Of TInput, TOutput)) As TOutput
        Return converter(InputData)
    End Function
End Structure

Module Program
    Sub Main()
        Dim cs As New ConverterStruct(Of String)("123")
        Dim res As Integer = cs.ConvertTo(Of Integer)(Function(s) Integer.Parse(s))
        __Check(CStr(res), "123")
    End Sub
End Module

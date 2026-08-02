' vybe-test: vb/vb_interface_default_methods_adv/test_vb_interface_generic_method
' origin: languages/vb/tests/vb/test_vb_interface_default_methods_adv.rs

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

Interface IConverter
    Function Convert(Of TInput, TOutput)(input As TInput) As TOutput
End Interface

Class StringConverter
    Implements IConverter
    Public Function Convert(Of TInput, TOutput)(input As TInput) As TOutput Implements IConverter.Convert
        Return CType(CObj(input.ToString()), TOutput)
    End Function
End Class

Module Program
    Sub Main()
        Dim c As IConverter = New StringConverter()
        Dim res As String = c.Convert(Of Integer, String)(100)
        __Check(CStr(res), "100")
    End Sub
End Module

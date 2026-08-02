' vybe-test: vb/vb_select_case_multiple_expressions/test_vb_select_case_type_checking_with_typeof
' origin: languages/vb/tests/vb/test_vb_select_case_multiple_expressions.rs

Module Program
    Sub Main()
        Dim obj As Object = "Hello"
        Select Case True
            Case TypeOf obj Is String
                Console.WriteLine("IsString")
            Case TypeOf obj Is Integer
                Console.WriteLine("IsInteger")
            Case Else
                Console.WriteLine("OtherType")
        End Select
    End Sub
End Module

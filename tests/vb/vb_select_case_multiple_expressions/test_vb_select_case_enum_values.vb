' vybe-test: vb/vb_select_case_multiple_expressions/test_vb_select_case_enum_values
' origin: languages/vb/tests/vb/test_vb_select_case_multiple_expressions.rs

Enum LogLevel
    Debug = 1
    Info = 2
    Error = 3
End Enum

Module Program
    Sub Main()
        Dim level = LogLevel.Error
        Select Case level
            Case LogLevel.Debug, LogLevel.Info
                Console.WriteLine("Non-Critical")
            Case LogLevel.Error
                Console.WriteLine("Critical")
        End Select
    End Sub
End Module

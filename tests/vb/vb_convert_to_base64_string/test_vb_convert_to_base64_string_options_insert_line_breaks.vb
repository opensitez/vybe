' vybe-test: vb/vb_convert_to_base64_string/test_vb_convert_to_base64_string_options_insert_line_breaks
' origin: languages/vb/tests/vb/test_vb_convert_to_base64_string.rs

Imports System

Module Program
    Sub Main()
        Dim largePayload As Byte() = New Byte(99) {}
        For i As Integer = 0 To 99 : largePayload(i) = CByte(i) : Next
        Dim b64Formatted = Convert.ToBase64String(largePayload, Base64FormattingOptions.InsertLineBreaks)
        Console.WriteLine(b64Formatted.Contains(vbCrLf) OrElse b64Formatted.Contains(vbLf))
    End Sub
End Module

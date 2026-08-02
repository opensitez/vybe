' vybe-test: vb/vb_do_loop_while_until_variants/test_vb_do_loop_with_try_catch_inside
' origin: languages/vb/tests/vb/test_vb_do_loop_while_until_variants.rs

Imports System

Module Program
    Sub Main()
        Dim attempts = 0
        Dim successCount = 0
        Do While attempts < 3
            attempts += 1
            Try
                If attempts = 2 Then Throw New Exception("Transient Error")
                successCount += 1
            Catch ex As Exception
                Console.WriteLine("Error Handled at Attempt " & attempts)
            End Try
        Loop
        Console.WriteLine(successCount)
    End Sub
End Module

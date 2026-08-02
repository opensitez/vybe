' vybe-test: vb/vb_divide_by_zero_exception_handling/test_vb_backslash_operator_integer_division_strict
' origin: languages/vb/tests/vb/test_vb_divide_by_zero_exception_handling.rs

Imports System

Module Program
    Sub Main()
        Try
            Dim a As Integer = 10
            Dim b As Integer = 0
            ' Visual Basic "\" operator performs strict integer division and throws!
            Dim res As Integer = a \ b
            Console.WriteLine(res)
        Catch ex As DivideByZeroException
            Console.WriteLine("Backslash Throws DivideByZeroException")
        End Try
    End Sub
End Module

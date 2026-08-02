' vybe-test: vb/vb_divide_by_zero_exception_handling/test_vb_decimal_division_by_zero_throws_divide_by_zero_exception
' origin: languages/vb/tests/vb/test_vb_divide_by_zero_exception_handling.rs

Imports System

Module Program
    Sub Main()
        Try
            Dim a As Decimal = 10.5D
            Dim b As Decimal = 0D
            Dim res As Decimal = a / b
            Console.WriteLine(res)
        Catch ex As DivideByZeroException
            Console.WriteLine("Decimal DivideByZeroException Caught")
        End Try
    End Sub
End Module

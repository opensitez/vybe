' vybe-test: vb/vb_divide_by_zero_exception_handling/test_vb_long_integer_division_by_zero
' origin: languages/vb/tests/vb/test_vb_divide_by_zero_exception_handling.rs

Imports System

Module Program
    Sub Main()
        Try
            Dim a As Long = 1000000000000L
            Dim b As Long = 0L
            Dim res As Long = a \ b
            Console.WriteLine(res)
        Catch ex As DivideByZeroException
            Console.WriteLine("Long DivideByZeroException Caught")
        End Try
    End Sub
End Module

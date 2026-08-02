' vybe-test: vb/vb_divide_by_zero_exception_handling/test_vb_short_integer_division_by_zero
' origin: languages/vb/tests/vb/test_vb_divide_by_zero_exception_handling.rs

Imports System

Module Program
    Sub Main()
        Try
            Dim a As Short = 100S
            Dim b As Short = 0S
            Dim res As Short = CShort(a \ b)
            Console.WriteLine(res)
        Catch ex As DivideByZeroException
            Console.WriteLine("Short DivideByZeroException Caught")
        End Try
    End Sub
End Module

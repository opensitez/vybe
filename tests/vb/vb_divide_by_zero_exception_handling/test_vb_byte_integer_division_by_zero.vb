' vybe-test: vb/vb_divide_by_zero_exception_handling/test_vb_byte_integer_division_by_zero
' origin: languages/vb/tests/vb/test_vb_divide_by_zero_exception_handling.rs

Imports System

Module Program
    Sub Main()
        Try
            Dim a As Byte = 255
            Dim b As Byte = 0
            Dim res As Byte = CByte(a \ b)
            Console.WriteLine(res)
        Catch ex As DivideByZeroException
            Console.WriteLine("Byte DivideByZeroException Caught")
        End Try
    End Sub
End Module

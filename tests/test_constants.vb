Module MainModule
    ' Public constant
    Public Const MAX_SIZE As Integer = 100
    ' Private constant
    Private Const APP_NAME As String = "My VB App"
    ' Constant with expression
    Const PI As Double = 3.14159

    Sub Main()
        Console.WriteLine("Testing Constants...")

        ' Use constants in code
        Dim size As Integer = MAX_SIZE
        If size = 100 Then
            Console.WriteLine("SUCCESS: Public constant access")
        Else
            Console.WriteLine("FAILURE: Public constant access")
        End If

        If APP_NAME = "My VB App" Then
            Console.WriteLine("SUCCESS: Private constant access")
        Else
            Console.WriteLine("FAILURE: Private constant access")
        End If

        ' Using constant in expression
        Dim area As Double = PI * 10 * 10
        ' Use approximate comparison for doubles if needed, but here PI is just a constant value
        If area = 314.159 Then
            Console.WriteLine("SUCCESS: Constant expression")
        Else
            Console.WriteLine("FAILURE: Constant expression. Got: " & CStr(area))
        End If

        Console.WriteLine("Constants Tests Completed")
    End Sub
End Module

' vybe-test: vb/vb_conditional_compilation/conditional_compilation
' origin: languages/vb/tests/vb/test_vb_conditional_compilation.rs

#Const DEBUG_MODE = True
#Const VERSION = 2

Module M
    Sub Main()
#If DEBUG_MODE Then
        Console.WriteLine("Debug On")
#Else
        Console.WriteLine("Debug Off")
#End If

#If VERSION = 1 Then
        Console.WriteLine("V1")
#ElseIf VERSION = 2 Then
        Console.WriteLine("V2")
#Else
        Console.WriteLine("Unknown")
#End If
    End Sub
End Module

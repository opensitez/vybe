' vybe-test: vb/vb_conditional_compilation_adv/conditional_compilation_adv
' origin: languages/vb/tests/vb/test_vb_conditional_compilation_adv.rs

#Const DEBUG = True

Module M
    Sub Main()
#If DEBUG Then
        Console.WriteLine("DebugMode")
#ElseIf RELEASE Then
        Console.WriteLine("ReleaseMode")
#Else
        Console.WriteLine("OtherMode")
#End If
    End Sub
End Module

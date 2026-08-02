' vybe-test: vb/vb_variables_and_consts/var_static_init_once
' origin: languages/vb/tests/vb/test_vb_variables_and_consts.rs

Module M
Function Init() As Integer
Console.WriteLine("Init")
Return 5
End Function
Sub Test()
Static x As Integer = Init()
x += 1
Console.WriteLine(x)
End Sub
Sub Main()
Test()
Test()
End Sub
End Module

' vybe-test: vb/vb_variables_and_consts/var_static_basic
' origin: languages/vb/tests/vb/test_vb_variables_and_consts.rs

Module M
Sub Test()
Static x As Integer = 0
x += 1
Console.WriteLine(x)
End Sub
Sub Main()
Test()
Test()
End Sub
End Module

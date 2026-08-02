' vybe-test: vb/vb_my_namespace/my_namespace_parsing
' origin: languages/vb/tests/vb/test_vb_my_namespace.rs

Module M
    Sub Main()
        ' My is a virtual namespace in VB.NET
        ' Usually includes My.Application, My.Computer, My.User
        
        ' Just checking compiler support for 'My' namespace 
        ' (availability of properties depends on the framework version)
        Dim b As Boolean = True
        If Not b Then
            Console.WriteLine(My.Computer.Name)
            Console.WriteLine(My.Application.Info.Title)
            Console.WriteLine(My.User.Name)
        End If
        Console.WriteLine("My Namespace Parsed")
    End Sub
End Module

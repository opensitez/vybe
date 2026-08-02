' vybe-test: vb/vb_module_semantics/module_semantics
' origin: languages/vb/tests/vb/test_vb_module_semantics.rs

Module GlobalConfig
    ' Modules are essentially NotInheritable classes with only Shared members
    Public Property AppName As String = "VybeApp"
    
    Public Sub PrintConfig()
        Console.WriteLine(AppName)
    End Sub
End Module

Module M
    Sub Main()
        ' Members of a module can be accessed without qualification
        PrintConfig()
        
        ' Or with qualification
        GlobalConfig.AppName = "VybeApp V2"
        GlobalConfig.PrintConfig()
    End Sub
End Module

' vybe-test: vb/vb_class_notinheritable/class_notoverridable_methods
' origin: languages/vb/tests/vb/test_vb_class_notinheritable.rs

Class BasePrinter
    Public Overridable Sub Print()
        Console.WriteLine("Base")
    End Sub
End Class

Class FastPrinter
    Inherits BasePrinter
    
    ' Seals the method from further overriding in derived classes
    Public NotOverridable Overrides Sub Print()
        Console.WriteLine("Fast")
    End Sub
End Class

Module M
    Sub Main()
        Dim fp As BasePrinter = New FastPrinter()
        fp.Print()
    End Sub
End Module

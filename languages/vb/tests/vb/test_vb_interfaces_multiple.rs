use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Interfaces (Multiple Inheritance and Implementation)
// ═══════════════════════════════════════════════════════════

#[test]
fn interface_multiple_implementation() {
    let out = run_vb(
        r#"
Interface IReader
    Sub Read()
End Interface

Interface IWriter
    Sub Write()
End Interface

Class FileIO
    Implements IReader, IWriter
    
    Public Sub Read() Implements IReader.Read
        Console.WriteLine("Reading")
    End Sub
    
    Public Sub Write() Implements IWriter.Write
        Console.WriteLine("Writing")
    End Sub
End Class

Module M
    Sub Main()
        Dim io As New FileIO()
        Dim r As IReader = io
        Dim w As IWriter = io
        
        r.Read()
        w.Write()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Reading", "Writing"]);
}

#[test]
fn interface_inheritance() {
    let out = run_vb(
        r#"
Interface IBase
    Sub MethodA()
End Interface

Interface IDerived
    Inherits IBase
    Sub MethodB()
End Interface

Class MyClass
    Implements IDerived
    
    Public Sub MethodA() Implements IDerived.MethodA
        Console.WriteLine("A")
    End Sub
    
    Public Sub MethodB() Implements IDerived.MethodB
        Console.WriteLine("B")
    End Sub
End Class

Module M
    Sub Main()
        Dim d As IDerived = New MyClass()
        d.MethodA()
        d.MethodB()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["A", "B"]);
}

#[test]
fn interface_implementing_same_method_from_multiple() {
    let out = run_vb(
        r#"
Interface IControl
    Sub Paint()
End Interface

Interface ISurface
    Sub Paint()
End Interface

Class Canvas
    Implements IControl, ISurface
    
    ' One method satisfies both interfaces
    Public Sub Paint() Implements IControl.Paint, ISurface.Paint
        Console.WriteLine("Painted both")
    End Sub
End Class

Module M
    Sub Main()
        Dim c As New Canvas()
        Dim c1 As IControl = c
        Dim c2 As ISurface = c
        
        c1.Paint()
        c2.Paint()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Painted both", "Painted both"]);
}

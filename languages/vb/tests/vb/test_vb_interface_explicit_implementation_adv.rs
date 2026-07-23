use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Interface Explicit Implementation & Multiple Implements
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_interface_multiple_same_name_members() {
    let src = r#"
Interface IControl
    Sub Paint()
End Interface

Interface ISurface
    Sub Paint()
End Interface

Class Canvas
    Implements IControl, ISurface

    Private Sub PaintControl() Implements IControl.Paint
        Console.WriteLine("Control Paint")
    End Sub

    Private Sub PaintSurface() Implements ISurface.Paint
        Console.WriteLine("Surface Paint")
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As New Canvas()
        Dim ctrl As IControl = c
        Dim surf As ISurface = c
        ctrl.Paint()
        surf.Paint()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Control Paint", "Surface Paint"]);
}

#[test]
fn test_vb_interface_member_mapping_different_name() {
    let src = r#"
Interface IWorker
    Sub DoWork()
End Interface

Class Employee
    Implements IWorker

    Public Sub ExecuteJob() Implements IWorker.DoWork
        Console.WriteLine("Executing Job")
    End Sub
End Class

Module Program
    Sub Main()
        Dim e As New Employee()
        e.ExecuteJob()
        Dim w As IWorker = e
        w.DoWork()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Executing Job", "Executing Job"]);
}

#[test]
fn test_vb_interface_implements_multiple_members_with_one_method() {
    let src = r#"
Interface IFoo
    Sub Process()
End Interface

Interface IBar
    Sub Process()
End Interface

Class Processor
    Implements IFoo, IBar

    Public Sub CommonProcess() Implements IFoo.Process, IBar.Process
        Console.WriteLine("Common Process")
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New Processor()
        Dim f As IFoo = p
        Dim b As IBar = p
        f.Process()
        b.Process()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Common Process", "Common Process"]);
}

#[test]
fn test_vb_interface_property_explicit_implementation() {
    let src = r#"
Interface INamed
    Property Name As String
End Interface

Class User
    Implements INamed
    Private _userName As String

    Private Property NameProp As String Implements INamed.Name
        Get
            Return _userName
        End Get
        Set(value As String)
            _userName = value
        End Set
    End Property

    Public Sub New(name As String)
        _userName = name
    End Sub
End Class

Module Program
    Sub Main()
        Dim u As New User("Bob")
        Dim n As INamed = u
        Console.WriteLine(n.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Bob"]);
}

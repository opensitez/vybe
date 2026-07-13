use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Interfaces (Basic Implementation)
// ═══════════════════════════════════════════════════════════

#[test]
fn interface_basic_implementation() {
    let out = run_vb(
        r#"
Interface IAnimal
    Sub Speak()
    Function GetName() As String
End Interface

Class Dog
    Implements IAnimal
    
    Public Sub Speak() Implements IAnimal.Speak
        Console.WriteLine("Woof")
    End Sub
    
    Public Function GetName() As String Implements IAnimal.GetName
        Return "Buddy"
    End Function
End Class

Module M
    Sub Main()
        Dim a As IAnimal = New Dog()
        a.Speak()
        Console.WriteLine(a.GetName())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Woof", "Buddy"]);
}

#[test]
fn interface_property_implementation() {
    let out = run_vb(
        r#"
Interface IVehicle
    Property Speed As Integer
End Interface

Class Car
    Implements IVehicle
    
    Private _speed As Integer
    Public Property Speed As Integer Implements IVehicle.Speed
        Get
            Return _speed
        End Get
        Set(value As Integer)
            _speed = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim v As IVehicle = New Car()
        v.Speed = 55
        Console.WriteLine(v.Speed)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["55"]);
}

#[test]
fn interface_event_implementation() {
    let out = run_vb(
        r#"
Interface IAlarm
    Event Triggered()
End Interface

Class SecuritySystem
    Implements IAlarm
    
    Public Event Triggered() Implements IAlarm.Triggered
    
    Public Sub SoundAlarm()
        RaiseEvent Triggered()
    End Sub
End Class

Module M
    Private WithEvents sys As SecuritySystem
    
    Private Sub sys_Triggered() Handles sys.Triggered
        Console.WriteLine("Alert!")
    End Sub
    
    Sub Main()
        sys = New SecuritySystem()
        sys.SoundAlarm()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Alert!"]);
}

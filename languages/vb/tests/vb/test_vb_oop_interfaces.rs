use super::helpers::run_vb;

#[test]
fn interface_basic_implementation() {
    assert_eq!(
        run_vb(
            "Interface I\nSub M()\nEnd Interface\nClass C\nImplements I\nPublic Sub M() Implements I.M\nConsole.WriteLine(\"C\")\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim c1 As I = New C()\nc1.M()\nEnd Sub\nEnd Module"
        ),
        vec!["C"]
    );
}
#[test]
fn interface_implicit_implementation_fails() {
    assert_eq!(
        run_vb(
            "Interface I\nSub M()\nEnd Interface\nClass C\nImplements I\n' Public Sub M() ' VB requires explicit Implements I.M\nEnd Class\nModule M\nSub Main()\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"
        ),
        vec!["Parsed"]
    );
}
#[test]
fn interface_explicit_renaming() {
    assert_eq!(
        run_vb(
            "Interface I\nSub M()\nEnd Interface\nClass C\nImplements I\nPublic Sub MyMethod() Implements I.M\nConsole.WriteLine(\"C\")\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim c1 As I = New C()\nc1.M()\nEnd Sub\nEnd Module"
        ),
        vec!["C"]
    );
} // Implements maps MyMethod to I.M
#[test]
fn interface_private_implementation() {
    assert_eq!(
        run_vb(
            "Interface I\nSub M()\nEnd Interface\nClass C\nImplements I\nPrivate Sub M() Implements I.M\nConsole.WriteLine(\"C\")\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim c1 As I = New C()\nc1.M()\nEnd Sub\nEnd Module"
        ),
        vec!["C"]
    );
} // Interface implementation can be private
#[test]
fn interface_property_implementation() {
    assert_eq!(
        run_vb(
            "Interface I\nProperty V As Integer\nEnd Interface\nClass C\nImplements I\nPublic Property V As Integer Implements I.V\nEnd Class\nModule M\nSub Main()\nDim c1 As I = New C()\nc1.V = 10\nConsole.WriteLine(c1.V)\nEnd Sub\nEnd Module"
        ),
        vec!["10"]
    );
}

#[test]
fn interface_multiple_interfaces() {
    assert_eq!(
        run_vb(
            "Interface I1\nSub M1()\nEnd Interface\nInterface I2\nSub M2()\nEnd Interface\nClass C\nImplements I1, I2\nPublic Sub M1() Implements I1.M1\nConsole.WriteLine(\"1\")\nEnd Sub\nPublic Sub M2() Implements I2.M2\nConsole.WriteLine(\"2\")\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nc1.M1()\nc1.M2()\nEnd Sub\nEnd Module"
        ),
        vec!["1", "2"]
    );
}
#[test]
fn interface_implement_multiple_methods_with_one() {
    assert_eq!(
        run_vb(
            "Interface I1\nSub M()\nEnd Interface\nInterface I2\nSub M()\nEnd Interface\nClass C\nImplements I1, I2\nPublic Sub M() Implements I1.M, I2.M\nConsole.WriteLine(\"C\")\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim c1 As I2 = New C()\nc1.M()\nEnd Sub\nEnd Module"
        ),
        vec!["C"]
    );
} // One method implements both
#[test]
fn interface_inheritance() {
    assert_eq!(
        run_vb(
            "Interface I1\nSub M1()\nEnd Interface\nInterface I2\nInherits I1\nSub M2()\nEnd Interface\nClass C\nImplements I2\nPublic Sub M1() Implements I2.M1\nConsole.WriteLine(\"1\")\nEnd Sub\nPublic Sub M2() Implements I2.M2\nConsole.WriteLine(\"2\")\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nc1.M1()\nEnd Sub\nEnd Module"
        ),
        vec!["1"]
    );
} // Inherited members are accessed via I2
#[test]
fn interface_shadowing_member() {
    assert_eq!(
        run_vb(
            "Interface I1\nSub M()\nEnd Interface\nInterface I2\nInherits I1\nShadows Sub M()\nEnd Interface\nClass C\nImplements I2\nPublic Sub M1() Implements I1.M\nConsole.WriteLine(\"1\")\nEnd Sub\nPublic Sub M2() Implements I2.M\nConsole.WriteLine(\"2\")\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim c1 As I1 = New C()\nc1.M()\nEnd Sub\nEnd Module"
        ),
        vec!["1"]
    );
}
#[test]
fn interface_missing_implementation_fails() {
    assert_eq!(
        run_vb(
            "Interface I\nSub M1()\nSub M2()\nEnd Interface\n' Class C: Implements I: Public Sub M1() Implements I.M1: End Sub ' Missing M2\nModule M\nSub Main()\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"
        ),
        vec!["Parsed"]
    );
}

#[test]
fn interface_event_implementation() {
    assert_eq!(
        run_vb(
            "Interface I\nEvent E()\nEnd Interface\nClass C\nImplements I\nPublic Event E() Implements I.E\nPublic Sub Raise()\nRaiseEvent E()\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1 IsNot Nothing)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn interface_default_property() {
    assert_eq!(
        run_vb(
            "Interface I\nDefault Property Item(i As Integer) As Integer\nEnd Interface\nClass C\nImplements I\nDefault Public Property Item(i As Integer) As Integer Implements I.Item\nGet\nReturn i\nEnd Get\nSet\nEnd Set\nEnd Property\nEnd Class\nModule M\nSub Main()\nDim c1 As I = New C()\nConsole.WriteLine(c1(5))\nEnd Sub\nEnd Module"
        ),
        vec!["5"]
    );
}
#[test]
fn interface_type_conversion_typeof() {
    assert_eq!(
        run_vb(
            "Interface I\nEnd Interface\nClass C\nImplements I\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(TypeOf c1 Is I)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn interface_type_conversion_directcast() {
    assert_eq!(
        run_vb(
            "Interface I\nEnd Interface\nClass C\nImplements I\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nDim i1 = DirectCast(c1, I)\nConsole.WriteLine(i1 IsNot Nothing)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn interface_type_conversion_trycast() {
    assert_eq!(
        run_vb(
            "Interface I\nEnd Interface\nClass C\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nDim i1 = TryCast(c1, I)\nConsole.WriteLine(i1 Is Nothing)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
} // TryCast to unimplemented interface returns Nothing

#[test]
fn interface_structure_implementation() {
    assert_eq!(
        run_vb(
            "Interface I\nFunction GetV() As Integer\nEnd Interface\nStructure S\nImplements I\nPublic Function GetV() As Integer Implements I.GetV\nReturn 42\nEnd Function\nEnd Structure\nModule M\nSub Main()\nDim s1 As I = New S()\nConsole.WriteLine(s1.GetV())\nEnd Sub\nEnd Module"
        ),
        vec!["42"]
    );
} // Boxing occurs here implicitly
#[test]
fn interface_polymorphism_list() {
    assert_eq!(
        run_vb(
            "Interface I\nFunction M() As String\nEnd Interface\nClass A\nImplements I\nPublic Function M() As String Implements I.M\nReturn \"A\"\nEnd Function\nEnd Class\nClass B\nImplements I\nPublic Function M() As String Implements I.M\nReturn \"B\"\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim arr() As I = {New A(), New B()}\nConsole.WriteLine(arr(0).M() & arr(1).M())\nEnd Sub\nEnd Module"
        ),
        vec!["AB"]
    );
}
#[test]
fn interface_signature_mismatch_fails() {
    assert_eq!(
        run_vb(
            "Interface I\nSub M(v As Integer)\nEnd Interface\nClass C\nImplements I\n' Public Sub M(v As String) Implements I.M ' Fails signature match\nEnd Class\nModule M\nSub Main()\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"
        ),
        vec!["Parsed"]
    );
}
#[test]
fn interface_abstract_base_class() {
    assert_eq!(
        run_vb(
            "Interface I\nSub M()\nEnd Interface\nMustInherit Class B\nImplements I\nPublic MustOverride Sub M() Implements I.M\nEnd Class\nClass C\nInherits B\nPublic Overrides Sub M()\nConsole.WriteLine(\"C\")\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim c1 As I = New C()\nc1.M()\nEnd Sub\nEnd Module"
        ),
        vec!["C"]
    );
} // Base delegates to MustOverride
#[test]
fn interface_implementing_inherited_interface_methods() {
    assert_eq!(
        run_vb(
            "Interface IBase\nSub Base()\nEnd Interface\nInterface IDerived\nInherits IBase\nSub Derived()\nEnd Interface\nClass C\nImplements IDerived\nPublic Sub Base() Implements IDerived.Base\nEnd Sub\nPublic Sub Derived() Implements IDerived.Derived\nConsole.WriteLine(\"D\")\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim c1 As IDerived = New C()\nc1.Derived()\nEnd Sub\nEnd Module"
        ),
        vec!["D"]
    );
}

#[test]
fn interface_overloading_methods() {
    assert_eq!(
        run_vb(
            "Interface I\nFunction M(v As Integer) As String\nFunction M(v As String) As String\nEnd Interface\nClass C\nImplements I\nPublic Function M(v As Integer) As String Implements I.M\nReturn \"Int\"\nEnd Function\nPublic Function M(v As String) As String Implements I.M\nReturn \"Str\"\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim c1 As I = New C()\nConsole.WriteLine(c1.M(5))\nEnd Sub\nEnd Module"
        ),
        vec!["Int"]
    );
}
#[test]
fn interface_implementation_with_byref() {
    assert_eq!(
        run_vb(
            "Interface I\nSub Mutate(ByRef v As Integer)\nEnd Interface\nClass C\nImplements I\nPublic Sub Mutate(ByRef v As Integer) Implements I.Mutate\nv = 10\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim c1 As I = New C()\nDim x = 1\nc1.Mutate(x)\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"
        ),
        vec!["10"]
    );
}
#[test]
fn interface_as_property_type() {
    assert_eq!(
        run_vb(
            "Interface I\nFunction M() As Integer\nEnd Interface\nClass C\nImplements I\nPublic Function M() As Integer Implements I.M\nReturn 42\nEnd Function\nEnd Class\nClass Container\nPublic Property Obj As I\nEnd Class\nModule M\nSub Main()\nDim cont As New Container()\ncont.Obj = New C()\nConsole.WriteLine(cont.Obj.M())\nEnd Sub\nEnd Module"
        ),
        vec!["42"]
    );
}
#[test]
fn interface_cannot_contain_fields() {
    assert_eq!(
        run_vb(
            "Interface I\n' Dim x As Integer ' Interfaces cannot contain fields\nEnd Interface\nModule M\nSub Main()\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"
        ),
        vec!["Parsed"]
    );
}
#[test]
fn interface_generic() {
    assert_eq!(
        run_vb(
            "Interface I(Of T)\nFunction GetV() As T\nEnd Interface\nClass C\nImplements I(Of Integer)\nPublic Function GetV() As Integer Implements I(Of Integer).GetV\nReturn 99\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim c1 As I(Of Integer) = New C()\nConsole.WriteLine(c1.GetV())\nEnd Sub\nEnd Module"
        ),
        vec!["99"]
    );
}

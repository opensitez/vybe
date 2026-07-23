use super::helpers::run_vb;

#[test]
fn access_public_class() {
    assert_eq!(
        run_vb(
            "Public Class C\nPublic V As Integer = 1\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.V)\nEnd Sub\nEnd Module"
        ),
        vec!["1"]
    );
}
#[test]
fn access_private_class() {
    assert_eq!(
        run_vb(
            "Private Class C\nPublic V As Integer = 1\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.V)\nEnd Sub\nEnd Module"
        ),
        vec!["1"]
    );
}
#[test]
fn access_friend_class() {
    assert_eq!(
        run_vb(
            "Friend Class C\nPublic V As Integer = 1\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.V)\nEnd Sub\nEnd Module"
        ),
        vec!["1"]
    );
}
#[test]
fn access_protected_class() {
    assert_eq!(
        run_vb(
            "Module M\n' Classes can only be protected if nested\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"
        ),
        vec!["Parsed"]
    );
}
#[test]
fn access_protected_friend_class() {
    assert_eq!(
        run_vb(
            "Module M\n' Classes can only be protected friend if nested\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"
        ),
        vec!["Parsed"]
    );
}

#[test]
fn access_public_method() {
    assert_eq!(
        run_vb(
            "Class C\nPublic Sub M1()\nConsole.WriteLine(\"Pub\")\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nc1.M1()\nEnd Sub\nEnd Module"
        ),
        vec!["Pub"]
    );
}
#[test]
fn access_private_method() {
    assert_eq!(
        run_vb(
            "Class C\nPrivate Sub M1()\nConsole.WriteLine(\"Priv\")\nEnd Sub\nPublic Sub CallM1()\nM1()\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nc1.CallM1()\nEnd Sub\nEnd Module"
        ),
        vec!["Priv"]
    );
}
#[test]
fn access_friend_method() {
    assert_eq!(
        run_vb(
            "Class C\nFriend Sub M1()\nConsole.WriteLine(\"Friend\")\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nc1.M1()\nEnd Sub\nEnd Module"
        ),
        vec!["Friend"]
    );
}
#[test]
fn access_protected_method() {
    assert_eq!(
        run_vb(
            "Class B\nProtected Sub M1()\nConsole.WriteLine(\"Prot\")\nEnd Sub\nEnd Class\nClass C\nInherits B\nPublic Sub CallM1()\nMyBase.M1()\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nc1.CallM1()\nEnd Sub\nEnd Module"
        ),
        vec!["Prot"]
    );
}
#[test]
fn access_protected_friend_method() {
    assert_eq!(
        run_vb(
            "Class B\nProtected Friend Sub M1()\nConsole.WriteLine(\"ProtFriend\")\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim b1 As New B()\nb1.M1()\nEnd Sub\nEnd Module"
        ),
        vec!["ProtFriend"]
    );
}

#[test]
fn access_public_field() {
    assert_eq!(
        run_vb(
            "Class C\nPublic V As Integer = 10\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.V)\nEnd Sub\nEnd Module"
        ),
        vec!["10"]
    );
}
#[test]
fn access_private_field() {
    assert_eq!(
        run_vb(
            "Class C\nPrivate V As Integer = 20\nPublic Function GetV() As Integer\nReturn V\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.GetV())\nEnd Sub\nEnd Module"
        ),
        vec!["20"]
    );
}
#[test]
fn access_friend_field() {
    assert_eq!(
        run_vb(
            "Class C\nFriend V As Integer = 30\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.V)\nEnd Sub\nEnd Module"
        ),
        vec!["30"]
    );
}
#[test]
fn access_protected_field() {
    assert_eq!(
        run_vb(
            "Class B\nProtected V As Integer = 40\nEnd Class\nClass C\nInherits B\nPublic Function GetV() As Integer\nReturn V\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.GetV())\nEnd Sub\nEnd Module"
        ),
        vec!["40"]
    );
}
#[test]
fn access_protected_friend_field() {
    assert_eq!(
        run_vb(
            "Class C\nProtected Friend V As Integer = 50\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.V)\nEnd Sub\nEnd Module"
        ),
        vec!["50"]
    );
}

#[test]
fn access_nested_private_class() {
    assert_eq!(
        run_vb(
            "Class Outer\nPrivate Class Inner\nPublic V As Integer = 60\nEnd Class\nPublic Function GetInnerV() As Integer\nDim i As New Inner()\nReturn i.V\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim o As New Outer()\nConsole.WriteLine(o.GetInnerV())\nEnd Sub\nEnd Module"
        ),
        vec!["60"]
    );
}
#[test]
fn access_nested_protected_class() {
    assert_eq!(
        run_vb(
            "Class Outer\nProtected Class Inner\nPublic V As Integer = 70\nEnd Class\nEnd Class\nClass Derived\nInherits Outer\nPublic Function GetV() As Integer\nDim i As New Inner()\nReturn i.V\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim d As New Derived()\nConsole.WriteLine(d.GetV())\nEnd Sub\nEnd Module"
        ),
        vec!["70"]
    );
}
#[test]
fn access_module_private_member() {
    assert_eq!(
        run_vb(
            "Module Data\nPrivate V As Integer = 80\nPublic Function GetV() As Integer\nReturn V\nEnd Function\nEnd Module\nModule M\nSub Main()\nConsole.WriteLine(Data.GetV())\nEnd Sub\nEnd Module"
        ),
        vec!["80"]
    );
}
#[test]
fn access_module_friend_member() {
    assert_eq!(
        run_vb(
            "Module Data\nFriend V As Integer = 90\nEnd Module\nModule M\nSub Main()\nConsole.WriteLine(Data.V)\nEnd Sub\nEnd Module"
        ),
        vec!["90"]
    );
}
#[test]
fn access_interface_public_only() {
    assert_eq!(
        run_vb(
            "Interface I\n' Interfaces can only have Public members\nSub Test()\nEnd Interface\nModule M\nSub Main()\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"
        ),
        vec!["Parsed"]
    );
}

#[test]
fn access_shadowing_private() {
    assert_eq!(
        run_vb(
            "Class B\nPrivate V As Integer = 1\nEnd Class\nClass C\nInherits B\nPublic Shadows V As Integer = 2\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.V)\nEnd Sub\nEnd Module"
        ),
        vec!["2"]
    );
}
#[test]
fn access_shadowing_protected() {
    assert_eq!(
        run_vb(
            "Class B\nProtected V As Integer = 1\nEnd Class\nClass C\nInherits B\nPrivate Shadows V As Integer = 2\nPublic Function GetV() As Integer\nReturn V\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.GetV())\nEnd Sub\nEnd Module"
        ),
        vec!["2"]
    );
}
#[test]
fn access_overriding_protected() {
    assert_eq!(
        run_vb(
            "Class B\nProtected Overridable Function M1() As String\nReturn \"B\"\nEnd Function\nEnd Class\nClass C\nInherits B\nProtected Overrides Function M1() As String\nReturn \"C\"\nEnd Function\nPublic Function CallM1() As String\nReturn M1()\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.CallM1())\nEnd Sub\nEnd Module"
        ),
        vec!["C"]
    );
}
#[test]
fn access_constructor_private() {
    assert_eq!(
        run_vb(
            "Class C\nPrivate Sub New()\nEnd Sub\nPublic Shared Function Create() As C\nReturn New C()\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim c1 = C.Create()\nConsole.WriteLine(c1 IsNot Nothing)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn access_constructor_protected() {
    assert_eq!(
        run_vb(
            "Class B\nProtected Sub New()\nEnd Sub\nEnd Class\nClass C\nInherits B\nPublic Sub New()\nMyBase.New()\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1 IsNot Nothing)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}

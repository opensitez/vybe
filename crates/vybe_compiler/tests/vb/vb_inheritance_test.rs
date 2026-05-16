//! Tests for VB.NET class inheritance as compiled by vybe_compiler_vb.
//!
//! Covers: Inherits, MyBase.New(), MyBase.Method(), method override,
//! field initialization, Shared methods, Property Get/Set, multi-level chains.

use super::helpers::run_vb;

// ---------------------------------------------------------------------------
// 1. Base class with field + method, create instance, call method
// ---------------------------------------------------------------------------
#[test]
fn t01_base_class_field_and_method() {
    let out = run_vb(r#"
Class Animal
    Public Name As String = "Unknown"

    Sub Speak()
        Console.WriteLine(Name)
    End Sub
End Class

Dim a As New Animal()
a.Speak()
"#);
    assert_eq!(out, vec!["Unknown"]);
}

// ---------------------------------------------------------------------------
// 2. Derived Inherits Base — inherited field accessible
// ---------------------------------------------------------------------------
#[test]
fn t02_inherited_field_accessible() {
    let out = run_vb(r#"
Class Animal
    Public Name As String = "Beast"
End Class

Class Dog
    Inherits Animal
End Class

Dim d As New Dog()
Console.WriteLine(d.Name)
"#);
    assert_eq!(out, vec!["Beast"]);
}

// ---------------------------------------------------------------------------
// 3. Derived Inherits Base — inherited method callable
// ---------------------------------------------------------------------------
#[test]
fn t03_inherited_method_callable() {
    let out = run_vb(r#"
Class Animal
    Public Name As String = "Cat"

    Function GetName() As String
        GetName = Name
    End Function
End Class

Class Cat
    Inherits Animal
End Class

Dim c As New Cat()
Console.WriteLine(c.GetName())
"#);
    assert_eq!(out, vec!["Cat"]);
}

// ---------------------------------------------------------------------------
// 4. Derived overrides method — override is used
// ---------------------------------------------------------------------------
#[test]
fn t04_derived_overrides_method() {
    let out = run_vb(r#"
Class Animal
    Sub Speak()
        Console.WriteLine("generic")
    End Sub
End Class

Class Dog
    Inherits Animal

    Sub Speak()
        Console.WriteLine("woof")
    End Sub
End Class

Dim d As New Dog()
d.Speak()
"#);
    assert_eq!(out, vec!["woof"]);
}

// ---------------------------------------------------------------------------
// 5. MyBase.New() in derived constructor — parent fields initialized
// ---------------------------------------------------------------------------
#[test]
fn t05_mybase_new_initializes_parent_fields() {
    let out = run_vb(r#"
Class Base
    Public X As Integer = 42
End Class

Class Child
    Inherits Base

    Sub New()
        MyBase.New()
    End Sub
End Class

Dim c As New Child()
Console.WriteLine(c.X)
"#);
    assert_eq!(out, vec!["42"]);
}

// ---------------------------------------------------------------------------
// 6. MyBase.New(args) with parameters
// ---------------------------------------------------------------------------
#[test]
fn t06_mybase_new_with_args() {
    let out = run_vb(r#"
Class Base
    Public Val As Integer

    Sub New(v As Integer)
        Val = v
    End Sub
End Class

Class Child
    Inherits Base

    Sub New(v As Integer)
        MyBase.New(v)
    End Sub
End Class

Dim c As New Child(99)
Console.WriteLine(c.Val)
"#);
    assert_eq!(out, vec!["99"]);
}

// ---------------------------------------------------------------------------
// 7. MyBase.Method() from override — calls parent version
// ---------------------------------------------------------------------------
#[test]
fn t07_mybase_method_calls_parent_version() {
    let out = run_vb(r#"
Class Base
    Sub Greet()
        Console.WriteLine("hello from base")
    End Sub
End Class

Class Child
    Inherits Base

    Sub Greet()
        MyBase.Greet()
        Console.WriteLine("hello from child")
    End Sub
End Class

Dim c As New Child()
c.Greet()
"#);
    assert_eq!(out, vec!["hello from base", "hello from child"]);
}

// ---------------------------------------------------------------------------
// 8. Three levels: GrandParent -> Parent -> Child
// ---------------------------------------------------------------------------
#[test]
fn t08_three_level_inheritance() {
    let out = run_vb(r#"
Class GrandParent
    Public A As String = "GP"
End Class

Class Parent
    Inherits GrandParent
    Public B As String = "P"
End Class

Class Child
    Inherits Parent
    Public C As String = "C"
End Class

Dim c As New Child()
Console.WriteLine(c.A)
Console.WriteLine(c.B)
Console.WriteLine(c.C)
"#);
    assert_eq!(out, vec!["GP", "P", "C"]);
}

// ---------------------------------------------------------------------------
// 9. Constructor chain through 3 levels
// ---------------------------------------------------------------------------
#[test]
fn t09_constructor_chain_three_levels() {
    let out = run_vb(r#"
Class Level1
    Public Tag As String

    Sub New()
        Tag = "L1"
        Console.WriteLine("Level1.New")
    End Sub
End Class

Class Level2
    Inherits Level1

    Sub New()
        MyBase.New()
        Tag = Tag & "-L2"
        Console.WriteLine("Level2.New")
    End Sub
End Class

Class Level3
    Inherits Level2

    Sub New()
        MyBase.New()
        Tag = Tag & "-L3"
        Console.WriteLine("Level3.New")
    End Sub
End Class

Dim x As New Level3()
Console.WriteLine(x.Tag)
"#);
    assert_eq!(out, vec!["Level1.New", "Level2.New", "Level3.New", "L1-L2-L3"]);
}

// ---------------------------------------------------------------------------
// 10. Derived adds new method, base methods still work
// ---------------------------------------------------------------------------
#[test]
fn t10_derived_adds_new_method() {
    let out = run_vb(r#"
Class Base
    Function Hello() As String
        Hello = "hi"
    End Function
End Class

Class Child
    Inherits Base

    Function World() As String
        World = "world"
    End Function
End Class

Dim c As New Child()
Console.WriteLine(c.Hello())
Console.WriteLine(c.World())
"#);
    assert_eq!(out, vec!["hi", "world"]);
}

// ---------------------------------------------------------------------------
// 11. Multiple methods: derived overrides one, inherits others
// ---------------------------------------------------------------------------
#[test]
fn t11_override_one_inherit_others() {
    let out = run_vb(r#"
Class Base
    Sub A()
        Console.WriteLine("base-A")
    End Sub

    Sub B()
        Console.WriteLine("base-B")
    End Sub
End Class

Class Child
    Inherits Base

    Sub A()
        Console.WriteLine("child-A")
    End Sub
End Class

Dim c As New Child()
c.A()
c.B()
"#);
    assert_eq!(out, vec!["child-A", "base-B"]);
}

// ---------------------------------------------------------------------------
// 12. No explicit Sub New in derived — parent auto-called
// ---------------------------------------------------------------------------
#[test]
fn t12_no_explicit_ctor_parent_auto_called() {
    let out = run_vb(r#"
Class Base
    Public Ready As String = "yes"

    Sub New()
        Console.WriteLine("base ctor")
    End Sub
End Class

Class Child
    Inherits Base
End Class

Dim c As New Child()
Console.WriteLine(c.Ready)
"#);
    assert_eq!(out, vec!["base ctor", "yes"]);
}

// ---------------------------------------------------------------------------
// 13. Base field set in Base.New, read from Derived method
// ---------------------------------------------------------------------------
#[test]
fn t13_base_field_set_in_ctor_read_from_derived() {
    let out = run_vb(r#"
Class Base
    Public Data As String

    Sub New()
        Data = "initialized"
    End Sub
End Class

Class Child
    Inherits Base

    Function GetData() As String
        GetData = Data
    End Function
End Class

Dim c As New Child()
Console.WriteLine(c.GetData())
"#);
    assert_eq!(out, vec!["initialized"]);
}

// ---------------------------------------------------------------------------
// 14. Derived field + base field both accessible
// ---------------------------------------------------------------------------
#[test]
fn t14_derived_and_base_fields_both_accessible() {
    let out = run_vb(r#"
Class Base
    Public X As Integer = 10
End Class

Class Child
    Inherits Base
    Public Y As Integer = 20
End Class

Dim c As New Child()
Console.WriteLine(c.X)
Console.WriteLine(c.Y)
"#);
    assert_eq!(out, vec!["10", "20"]);
}

// ---------------------------------------------------------------------------
// 15. Override calls MyBase then adds logic
// ---------------------------------------------------------------------------
#[test]
fn t15_override_calls_mybase_then_adds_logic() {
    let out = run_vb(r#"
Class Base
    Function Compute() As String
        Compute = "base"
    End Function
End Class

Class Child
    Inherits Base

    Function Compute() As String
        Dim b As String = MyBase.Compute()
        Compute = b & "+child"
    End Function
End Class

Dim c As New Child()
Console.WriteLine(c.Compute())
"#);
    assert_eq!(out, vec!["base+child"]);
}

// ---------------------------------------------------------------------------
// 16. Two different derived classes from same base — independent
// ---------------------------------------------------------------------------
#[test]
fn t16_two_derived_classes_independent() {
    let out = run_vb(r#"
Class Base
    Public Val As Integer = 0
End Class

Class ChildA
    Inherits Base

    Sub New()
        MyBase.New()
        Val = 1
    End Sub
End Class

Class ChildB
    Inherits Base

    Sub New()
        MyBase.New()
        Val = 2
    End Sub
End Class

Dim a As New ChildA()
Dim b As New ChildB()
Console.WriteLine(a.Val)
Console.WriteLine(b.Val)
"#);
    assert_eq!(out, vec!["1", "2"]);
}

// ---------------------------------------------------------------------------
// 17. Shared method on base — callable via ClassName.Method
// ---------------------------------------------------------------------------
#[test]
fn t17_shared_method_on_base() {
    let out = run_vb(r#"
Class MathHelper
    Shared Function Double(x As Integer) As Integer
        Double = x * 2
    End Function
End Class

Console.WriteLine(MathHelper.Double(5))
"#);
    assert_eq!(out, vec!["10"]);
}

// ---------------------------------------------------------------------------
// 18. Shared method inherited by derived
// ---------------------------------------------------------------------------
#[test]
fn t18_shared_method_inherited_by_derived() {
    let out = run_vb(r#"
Class Base
    Shared Function StaticHello() As String
        StaticHello = "hello"
    End Function
End Class

Class Child
    Inherits Base
End Class

Console.WriteLine(Child.StaticHello())
"#);
    assert_eq!(out, vec!["hello"]);
}

// ---------------------------------------------------------------------------
// 19. Property Get/Set on base class
// ---------------------------------------------------------------------------
#[test]
fn t19_property_get_set_on_base() {
    let out = run_vb(r#"
Class Person
    Private _name As String = ""

    Property Name() As String
        Get
            Name = _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property
End Class

Dim p As New Person()
p.Name = "Alice"
Console.WriteLine(p.Name)
"#);
    assert_eq!(out, vec!["Alice"]);
}

// ---------------------------------------------------------------------------
// 20. Derived accesses base Property
// ---------------------------------------------------------------------------
#[test]
fn t20_derived_accesses_base_property() {
    let out = run_vb(r#"
Class Base
    Private _val As Integer = 0

    Property Val() As Integer
        Get
            Val = _val
        End Get
        Set(value As Integer)
            _val = value
        End Set
    End Property
End Class

Class Child
    Inherits Base
End Class

Dim c As New Child()
c.Val = 77
Console.WriteLine(c.Val)
"#);
    assert_eq!(out, vec!["77"]);
}

//! Inheritance, virtual dispatch, base calls, and method hiding.
use super::helpers::run_csharp;

#[test]
fn derived_class_inherits_public_method_from_base() {
    assert_eq!(
        run_csharp(
            r#"class Base { public string Hello() => "hello"; }
class Derived : Base { }
Console.WriteLine(new Derived().Hello());"#
        ),
        &["hello"]
    );
}

#[test]
fn override_replaces_virtual_method_via_base_reference() {
    assert_eq!(
        run_csharp(
            r#"class Animal { public virtual string Sound() => "..."; }
class Dog : Animal { public override string Sound() => "woof"; }
Animal a = new Dog();
Console.WriteLine(a.Sound());"#
        ),
        &["woof"]
    );
}

#[test]
fn base_call_delegates_to_parent_implementation() {
    assert_eq!(
        run_csharp(
            r#"class A { public virtual string Greet() => "Hello"; }
class B : A { public override string Greet() => base.Greet() + " World"; }
Console.WriteLine(new B().Greet());"#
        ),
        &["Hello World"]
    );
}

#[test]
fn method_hiding_with_new_keyword_selects_by_static_type() {
    assert_eq!(
        run_csharp(
            r#"class Parent { public string Name() => "Parent"; }
class Child : Parent { public new string Name() => "Child"; }
Parent p = new Child();
Console.WriteLine(p.Name());"#
        ),
        &["Parent"]
    );
}

#[test]
fn constructor_chaining_with_base_passes_args_to_parent() {
    assert_eq!(
        run_csharp(
            r#"class Shape { public string Color; public Shape(string c) { Color = c; } }
class Box : Shape { public Box(string c) : base(c) { } }
Console.WriteLine(new Box("red").Color);"#
        ),
        &["red"]
    );
}

#[test]
fn abstract_class_provides_partial_implementation() {
    assert_eq!(
        run_csharp(
            r#"abstract class Base {
    public abstract int Value();
    public int Double() => Value() * 2;
}
class Impl : Base { public override int Value() => 5; }
Console.WriteLine(new Impl().Double());"#
        ),
        &["10"]
    );
}

#[test]
fn derived_class_can_have_additional_members() {
    assert_eq!(
        run_csharp(
            r#"class Vehicle { public int Wheels = 4; }
class Bike : Vehicle { public bool HasKickstand = true; }
var bike = new Bike();
Console.WriteLine(bike.Wheels);
Console.WriteLine(bike.HasKickstand);"#
        ),
        &["4", "True"]
    );
}

#[test]
fn is_operator_checks_runtime_type_in_hierarchy() {
    assert_eq!(
        run_csharp(
            r#"class A { }
class B : A { }
object obj = new B();
Console.WriteLine(obj is A);
Console.WriteLine(obj is B);"#
        ),
        &["True", "True"]
    );
}

#[test]
fn cast_to_base_succeeds_from_derived_instance() {
    assert_eq!(
        run_csharp(
            r#"class Base { public int X = 1; }
class Derived : Base { public int Y = 2; }
Base b = (Base)new Derived();
Console.WriteLine(b.X);"#
        ),
        &["1"]
    );
}

#[test]
fn protected_member_is_accessible_in_derived_class() {
    assert_eq!(
        run_csharp(
            r#"class Base { protected int Secret = 42; }
class Child : Base { public int Get() => Secret; }
Console.WriteLine(new Child().Get());"#
        ),
        &["42"]
    );
}

#[test]
fn object_tostring_is_overridable_for_custom_display() {
    assert_eq!(
        run_csharp(
            r#"class Point { public int X,Y; public override string ToString() => $"({X},{Y})"; }
Console.WriteLine(new Point { X=1, Y=2 });"#
        ),
        &["(1,2)"]
    );
}

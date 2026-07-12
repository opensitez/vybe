//! Generic type constraints: `where T : class`, `struct`, `new()`, interface, base type.
use super::helpers::run_csharp;

#[test]
fn new_constraint_allows_default_constructor_call_inside_generic_method() {
    assert_eq!(
        run_csharp(
            r#"
T Create<T>() where T : new() => new T();
class Widget { public int Value = 42; }
var w = Create<Widget>();
Console.WriteLine(w.Value);
"#
        ),
        &["42"]
    );
}

#[test]
fn class_constraint_allows_null_assignment_to_type_parameter() {
    assert_eq!(
        run_csharp(
            r#"
T AsNull<T>() where T : class => null;
Console.WriteLine(AsNull<string>() == null);
"#
        ),
        &["True"]
    );
}

#[test]
fn struct_constraint_produces_non_null_default_value() {
    assert_eq!(
        run_csharp(
            r#"
T Zero<T>() where T : struct => default;
Console.WriteLine(Zero<int>());
"#
        ),
        &["0"]
    );
}

#[test]
fn interface_constraint_enforces_method_availability_at_compile_time() {
    assert_eq!(
        run_csharp(
            r#"
interface ILabel { string Label(); }
class Tag : ILabel { public string Label() => "tag"; }
string Get<T>(T t) where T : ILabel => t.Label();
Console.WriteLine(Get(new Tag()));
"#
        ),
        &["tag"]
    );
}

#[test]
fn base_type_constraint_allows_method_call_on_constrained_parameter() {
    assert_eq!(
        run_csharp(
            r#"
class Animal { public virtual string Sound() => "..."; }
class Dog : Animal { public override string Sound() => "woof"; }
string Speak<T>(T t) where T : Animal => t.Sound();
Console.WriteLine(Speak(new Dog()));
"#
        ),
        &["woof"]
    );
}

#[test]
fn multiple_constraints_combine_with_comma_syntax() {
    assert_eq!(
        run_csharp(
            r#"
interface IName { string Name(); }
T Make<T>() where T : IName, new() => new T();
class Item : IName { public string Name() => "item"; }
Console.WriteLine(Make<Item>().Name());
"#
        ),
        &["item"]
    );
}

//! Virtual method dispatch, `sealed` overrides, and `new` method hiding —
//! distinct polymorphism contracts from ordinary method calls.
use super::helpers::run_csharp;

#[test]
fn virtual_call_through_base_reference_uses_most_derived_override() {
    assert_eq!(
        run_csharp(
            r#"
class Animal {
    public virtual string Speak() { return "..."; }
}
class Dog : Animal {
    public override string Speak() { return "woof"; }
}
Animal pet = new Dog();
Console.WriteLine(pet.Speak());
"#
        ),
        &["woof"]
    );
}

#[test]
fn sealed_override_prevents_further_overriding_in_grandchild() {
    assert_eq!(
        run_csharp(
            r#"
class Base {
    public virtual string Tag() { return "base"; }
}
class Middle : Base {
    public sealed override string Tag() { return "middle"; }
}
class Leaf : Middle {
    public override string Tag() { return "leaf"; }
}
Base item = new Leaf();
Console.WriteLine(item.Tag());
"#
        ),
        &["middle"]
    );
}

#[test]
fn method_hiding_with_new_keyword_does_not_change_base_reference_dispatch() {
    assert_eq!(
        run_csharp(
            r#"
class Base {
    public string Name() { return "base"; }
}
class Derived : Base {
    public new string Name() { return "derived"; }
}
Base reference = new Derived();
Derived concrete = new Derived();
Console.WriteLine(reference.Name());
Console.WriteLine(concrete.Name());
"#
        ),
        &["base", "derived"]
    );
}

#[test]
fn base_keyword_invokes_parent_implementation_from_override() {
    assert_eq!(
        run_csharp(
            r#"
class Counter {
    public virtual int Next() { return 1; }
}
class StepCounter : Counter {
    public override int Next() { return base.Next() + 2; }
}
Console.WriteLine(new StepCounter().Next());
"#
        ),
        &["3"]
    );
}

#[test]
fn virtual_property_getter_dispatches_to_derived_accessor() {
    assert_eq!(
        run_csharp(
            r#"
class Shape {
    public virtual int Sides { get { return 0; } }
}
class Triangle : Shape {
    public override int Sides { get { return 3; } }
}
Shape shape = new Triangle();
Console.WriteLine(shape.Sides);
"#
        ),
        &["3"]
    );
}

#[test]
fn abstract_method_must_be_implemented_by_concrete_derived_class() {
    assert_eq!(
        run_csharp(
            r#"
abstract class Parser {
    public abstract string Parse(string input);
}
class EchoParser : Parser {
    public override string Parse(string input) { return input.Trim(); }
}
Parser parser = new EchoParser();
Console.WriteLine(parser.Parse("  hi  "));
"#
        ),
        &["hi"]
    );
}

#[test]
fn calling_virtual_from_constructor_uses_current_type_override() {
    assert_eq!(
        run_csharp(
            r#"
class Base {
    public Base() { Console.WriteLine(Describe()); }
    public virtual string Describe() { return "base"; }
}
class Derived : Base {
    string label = "derived";
    public override string Describe() { return label; }
}
new Derived();
"#
        ),
        &[""]
    );
}

#[test]
fn non_virtual_method_call_ignores_derived_redeclaration_without_new() {
    assert_eq!(
        run_csharp(
            r#"
class Printer {
    public string Format(int value) { return "p:" + value; }
}
class FancyPrinter : Printer {
    public string Format(int value) { return "f:" + value; }
}
Printer tool = new FancyPrinter();
Console.WriteLine(tool.Format(7));
"#
        ),
        &["p:7"]
    );
}

#[test]
fn interface_method_implementation_is_invoked_through_interface_reference() {
    assert_eq!(
        run_csharp(
            r#"
interface IFormatter {
    string Format(int value);
}
class DecimalFormatter : IFormatter {
    public string Format(int value) { return value.ToString("D3"); }
}
IFormatter formatter = new DecimalFormatter();
Console.WriteLine(formatter.Format(4));
"#
        ),
        &["004"]
    );
}

#[test]
fn explicit_interface_implementation_not_visible_through_class_reference() {
    assert_eq!(
        run_csharp(
            r#"
interface IWorker {
    string Work();
}
class Person : IWorker {
    string IWorker.Work() { return "hidden"; }
    public string Work() { return "public"; }
}
Person person = new Person();
Console.WriteLine(person.Work());
Console.WriteLine(((IWorker)person).Work());
"#
        ),
        &["public", "hidden"]
    );
}

#[test]
fn chained_base_constructor_initializes_before_derived_fields() {
    assert_eq!(
        run_csharp(
            r#"
class Base {
    protected string token;
    public Base(string token) { this.token = token; }
}
class Child : Base {
    public string Label;
    public Child(string token, string label) : base(token) { Label = label; }
    public string Read() { return token + ":" + Label; }
}
Console.WriteLine(new Child("id", "name").Read());
"#
        ),
        &["id:name"]
    );
}

#[test]
fn this_constructor_chain_reuses_sibling_constructor_logic() {
    assert_eq!(
        run_csharp(
            r#"
class Pair {
    public int First;
    public int Second;
    public Pair(int value) : this(value, value) { }
    public Pair(int first, int second) { First = first; Second = second; }
}
var pair = new Pair(9);
Console.WriteLine(pair.First);
Console.WriteLine(pair.Second);
"#
        ),
        &["9", "9"]
    );
}

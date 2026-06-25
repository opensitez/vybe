//! Abstract class patterns: template method, abstract property, partial implementation.
use super::helpers::run_csharp;

#[test]
fn template_method_calls_abstract_hook_overridden_by_subclass() {
    assert_eq!(
        run_csharp(r#"abstract class Report{
    protected abstract string Header();
    protected abstract string Body();
    public string Generate()=>Header()+"\n"+Body();
}
class HtmlReport:Report{
    protected override string Header()=>"<html>";
    protected override string Body()=>"<body></body>";
}
Console.WriteLine(new HtmlReport().Generate());"#),
        &["<html>", "<body></body>"]
    );
}

#[test]
fn abstract_property_overridden_in_concrete_class() {
    assert_eq!(
        run_csharp(r#"abstract class Shape{public abstract double Area;}
class Square:Shape{public double Side;public override double Area=>Side*Side;}
Shape s=new Square{Side=4};
Console.WriteLine(s.Area);"#),
        &["16"]
    );
}

#[test]
fn abstract_class_can_have_concrete_methods_used_by_subclass() {
    assert_eq!(
        run_csharp(r#"abstract class Animal{
    public abstract string Sound();
    public string Speak()=>$"I say {Sound()}";
}
class Cat:Animal{public override string Sound()=>"meow";}
Console.WriteLine(new Cat().Speak());"#),
        &["I say meow"]
    );
}

#[test]
fn abstract_class_with_constructor_initialized_by_derived() {
    assert_eq!(
        run_csharp(r#"abstract class Named{public string Name;public Named(string n){Name=n;}}
class Tag:Named{public Tag(string n):base(n){}}
Console.WriteLine(new Tag("admin").Name);"#),
        &["admin"]
    );
}

#[test]
fn abstract_class_holding_state_shared_with_subclass() {
    assert_eq!(
        run_csharp(r#"abstract class Counter{
    protected int Count;
    public abstract void Increment();
    public int Value=>Count;
}
class By2:Counter{public override void Increment(){Count+=2;}}
var c=new By2(); c.Increment(); c.Increment();
Console.WriteLine(c.Value);"#),
        &["4"]
    );
}

#[test]
fn derived_abstract_class_can_leave_some_methods_unimplemented() {
    assert_eq!(
        run_csharp(r#"abstract class A{public abstract int X();public abstract int Y();}
abstract class B:A{public override int X()=>1;}
class C:B{public override int Y()=>2;}
var c=new C();
Console.WriteLine(c.X()); Console.WriteLine(c.Y());"#),
        &["1", "2"]
    );
}

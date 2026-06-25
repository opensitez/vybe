//! Primary constructors for classes and structs (C# 12).
use super::helpers::run_csharp;

#[test]
fn primary_constructor_parameters_available_in_methods() {
    assert_eq!(
        run_csharp(r#"class Counter(int start){
    int current=start;
    public int Next()=>++current;
    public int Value=>current;
}
var c=new Counter(10);
c.Next(); c.Next();
Console.WriteLine(c.Value);"#),
        &["12"]
    );
}

#[test]
fn primary_constructor_captures_dependency_for_methods() {
    assert_eq!(
        run_csharp(r#"class Greeter(string prefix){
    public string Greet(string name)=>$"{prefix} {name}";
}
Console.WriteLine(new Greeter("Hello").Greet("World"));"#),
        &["Hello World"]
    );
}

#[test]
fn primary_constructor_struct_exposes_parameter_as_field_via_manual_assignment() {
    assert_eq!(
        run_csharp(r#"struct Point(int x,int y){
    public int X=x;
    public int Y=y;
}
var p=new Point(3,4);
Console.WriteLine(p.X); Console.WriteLine(p.Y);"#),
        &["3", "4"]
    );
}

#[test]
fn primary_constructor_with_base_call_passes_args_up() {
    assert_eq!(
        run_csharp(r#"class Animal(string name){public string Name=>name;}
class Dog(string name,string breed):Animal(name){public string Breed=>breed;}
var d=new Dog("Rex","Lab");
Console.WriteLine(d.Name); Console.WriteLine(d.Breed);"#),
        &["Rex", "Lab"]
    );
}

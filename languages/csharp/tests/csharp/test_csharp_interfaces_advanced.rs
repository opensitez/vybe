//! Interface with default methods, static members, and explicit multiple implementation.
use super::helpers::run_csharp;

#[test]
fn default_interface_method_provides_fallback_implementation() {
    assert_eq!(
        run_csharp(
            r#"interface IGreeter{
    string Name();
    string Greet()=>"Hello "+Name();
}
class Alice:IGreeter{public string Name()=>"Alice";}
IGreeter g=new Alice();
Console.WriteLine(g.Greet());"#
        ),
        &["Hello Alice"]
    );
}

#[test]
fn overriding_default_interface_method_replaces_fallback() {
    assert_eq!(
        run_csharp(
            r#"interface IGreeter{
    string Name();
    string Greet()=>"Hello "+Name();
}
class Bob:IGreeter{
    public string Name()=>"Bob";
    public string Greet()=>"Hi "+Name()+"!";
}
IGreeter g=new Bob();
Console.WriteLine(g.Greet());"#
        ),
        &["Hi Bob!"]
    );
}

#[test]
fn class_implements_multiple_interfaces_with_different_methods() {
    assert_eq!(
        run_csharp(
            r#"interface IRead{string Read();}
interface IWrite{void Write(string v);}
class Buffer:IRead,IWrite{
    string _val="";
    public string Read()=>_val;
    public void Write(string v){_val=v;}
}
var b=new Buffer();
((IWrite)b).Write("data");
Console.WriteLine(((IRead)b).Read());"#
        ),
        &["data"]
    );
}

#[test]
fn interface_type_hierarchy_allows_derived_assignment() {
    assert_eq!(
        run_csharp(
            r#"interface IAnimal{string Kind();}
interface IPet:IAnimal{string Name();}
class Dog:IPet{public string Kind()=>"dog"; public string Name()=>"Rex";}
IPet pet=new Dog();
IAnimal animal=pet;
Console.WriteLine(animal.Kind());
Console.WriteLine(pet.Name());"#
        ),
        &["dog", "Rex"]
    );
}

#[test]
fn interface_implemented_by_struct_and_used_through_ref() {
    assert_eq!(
        run_csharp(
            r#"interface IArea{double Area();}
struct Rect:IArea{public double W,H; public double Area()=>W*H;}
IArea a=new Rect{W=3,H=4};
Console.WriteLine(a.Area());"#
        ),
        &["12"]
    );
}

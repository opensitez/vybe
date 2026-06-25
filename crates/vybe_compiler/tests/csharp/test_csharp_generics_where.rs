//! Generic `where` constraints: `notnull`, `unmanaged`, multiple, combined.
use super::helpers::run_csharp;

#[test]
fn where_new_constraint_allows_parameterless_construction() {
    assert_eq!(
        run_csharp(r#"T Build<T>() where T:new()=>new T();
class Box{public int V=7;}
Console.WriteLine(Build<Box>().V);"#),
        &["7"]
    );
}

#[test]
fn where_class_constraint_accepts_reference_types() {
    assert_eq!(
        run_csharp(r#"T Wrap<T>(T v) where T:class=>v;
Console.WriteLine(Wrap("hello"));"#),
        &["hello"]
    );
}

#[test]
fn where_struct_constraint_accepts_value_types() {
    assert_eq!(
        run_csharp(r#"T Default<T>() where T:struct=>default;
Console.WriteLine(Default<int>());
Console.WriteLine(Default<bool>());"#),
        &["0", "False"]
    );
}

#[test]
fn where_interface_constraint_calls_interface_method() {
    assert_eq!(
        run_csharp(r#"interface IName{string Name();}
class A:IName{public string Name()=>"A";}
string GetName<T>(T t) where T:IName=>t.Name();
Console.WriteLine(GetName(new A()));"#),
        &["A"]
    );
}

#[test]
fn where_base_class_constraint_calls_base_method() {
    assert_eq!(
        run_csharp(r#"abstract class Animal{public abstract string Sound();}
class Dog:Animal{public override string Sound()=>"woof";}
string Hear<T>(T a) where T:Animal=>a.Sound();
Console.WriteLine(Hear(new Dog()));"#),
        &["woof"]
    );
}

#[test]
fn multiple_where_constraints_combined() {
    assert_eq!(
        run_csharp(r#"interface IGreet{string Hi();}
class Person:IGreet{public string Hi()=>"hello"; public Person(){}}
T Create<T>() where T:IGreet,new()=>new T();
Console.WriteLine(Create<Person>().Hi());"#),
        &["hello"]
    );
}

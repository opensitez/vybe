//! Polymorphism: virtual dispatch, covariant returns, hiding with `new`, casting.
use super::helpers::run_csharp;

#[test]
fn virtual_method_dispatches_to_most_derived_override() {
    assert_eq!(
        run_csharp(
            r#"class Base{public virtual string Speak()=>"base";}
class Derived:Base{public override string Speak()=>"derived";}
Base obj=new Derived();
Console.WriteLine(obj.Speak());"#
        ),
        &["derived"]
    );
}

#[test]
fn method_hiding_with_new_does_not_override_base_dispatch() {
    assert_eq!(
        run_csharp(
            r#"class Base{public virtual string Speak()=>"base";}
class Derived:Base{public new string Speak()=>"hidden";}
Base obj=new Derived();
Console.WriteLine(obj.Speak());"#
        ),
        &["base"]
    );
}

#[test]
fn is_operator_succeeds_for_derived_held_as_base() {
    assert_eq!(
        run_csharp(
            r#"class Animal{} class Dog:Animal{}
Animal a=new Dog();
Console.WriteLine(a is Dog); Console.WriteLine(a is Animal);"#
        ),
        &["True", "True"]
    );
}

#[test]
fn as_operator_returns_null_for_incompatible_cast() {
    assert_eq!(
        run_csharp(
            r#"class A{} class B{}
object o=new A();
Console.WriteLine(o as B==null);"#
        ),
        &["True"]
    );
}

#[test]
fn polymorphic_list_iterates_dispatching_to_each_type() {
    assert_eq!(
        run_csharp(
            r#"abstract class Shape{public abstract int Size();}
class Square:Shape{public override int Size()=>4;}
class Triangle:Shape{public override int Size()=>3;}
var shapes=new System.Collections.Generic.List<Shape>{new Square(),new Triangle(),new Square()};
int sum=0; foreach(var s in shapes) sum+=s.Size();
Console.WriteLine(sum);"#
        ),
        &["11"]
    );
}

#[test]
fn direct_cast_throws_invalid_cast_for_unrelated_type() {
    assert_eq!(
        run_csharp(
            r#"string r="";
try{object o="hello"; int n=(int)o;}
catch(System.InvalidCastException){r="bad cast";}
Console.WriteLine(r);"#
        ),
        &["bad cast"]
    );
}

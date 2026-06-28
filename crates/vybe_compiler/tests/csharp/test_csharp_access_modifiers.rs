//! Access modifiers: public, private, protected, internal, protected internal, private protected.
use super::helpers::run_csharp;

#[test]
fn private_field_only_accessible_within_declaring_class() {
    assert_eq!(
        run_csharp(
            r#"class Safe{private int secret=42; public int Get()=>secret;}
Console.WriteLine(new Safe().Get());"#
        ),
        &["42"]
    );
}

#[test]
fn protected_field_accessible_in_subclass_method() {
    assert_eq!(
        run_csharp(
            r#"class A{protected int Value=7;}
class B:A{public int Read()=>Value;}
Console.WriteLine(new B().Read());"#
        ),
        &["7"]
    );
}

#[test]
fn internal_member_accessible_within_same_assembly() {
    assert_eq!(
        run_csharp(
            r#"class Library{internal string Tag="v1";}
Console.WriteLine(new Library().Tag);"#
        ),
        &["v1"]
    );
}

#[test]
fn public_method_callable_from_any_scope() {
    assert_eq!(
        run_csharp(
            r#"class Service{public string Name()=>"svc";}
Console.WriteLine(new Service().Name());"#
        ),
        &["svc"]
    );
}

#[test]
fn private_setter_means_field_read_only_from_outside() {
    assert_eq!(
        run_csharp(
            r#"class Counter{
    public int Count{get;private set;}
    public void Tick(){Count++;}
}
var c=new Counter(); c.Tick(); c.Tick();
Console.WriteLine(c.Count);"#
        ),
        &["2"]
    );
}

#[test]
fn sealed_class_is_not_further_derivable_but_usable() {
    assert_eq!(
        run_csharp(
            r#"sealed class Final{public int Value=99;}
Console.WriteLine(new Final().Value);"#
        ),
        &["99"]
    );
}

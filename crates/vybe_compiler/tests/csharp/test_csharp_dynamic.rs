//! `dynamic` keyword: late binding, `ExpandoObject`, and duck typing.
use super::helpers::run_csharp;

#[test]
fn dynamic_variable_holds_int_and_responds_to_arithmetic() {
    assert_eq!(
        run_csharp(r#"dynamic x=5; x+=3;
Console.WriteLine(x);"#),
        &["8"]
    );
}

#[test]
fn dynamic_variable_reassigned_to_different_type() {
    assert_eq!(
        run_csharp(r#"dynamic v=42;
Console.WriteLine(v.GetType().Name);
v="hello";
Console.WriteLine(v.GetType().Name);"#),
        &["Int32", "String"]
    );
}

#[test]
fn dynamic_expando_object_accepts_arbitrary_properties() {
    assert_eq!(
        run_csharp(r#"dynamic obj=new System.Dynamic.ExpandoObject();
obj.Name="Alice";
obj.Age=30;
Console.WriteLine(obj.Name); Console.WriteLine(obj.Age);"#),
        &["Alice", "30"]
    );
}

#[test]
fn dynamic_method_call_dispatched_at_runtime() {
    assert_eq!(
        run_csharp(r#"object o="hello";
dynamic d=o;
Console.WriteLine(d.ToUpper());"#),
        &["HELLO"]
    );
}

#[test]
fn dynamic_dictionary_access_via_expando() {
    assert_eq!(
        run_csharp(r#"dynamic e=new System.Dynamic.ExpandoObject();
var dict=(System.Collections.Generic.IDictionary<string,object>)e;
dict["x"]=99;
Console.WriteLine(e.x);"#),
        &["99"]
    );
}

//! Reflection: assembly metadata, `MemberInfo`, `TypeInfo`, `IsAssignableFrom`.
use super::helpers::run_csharp;

#[test]
fn type_is_assignable_from_derived_class() {
    assert_eq!(
        run_csharp(
            r#"class A{} class B:A{}
Console.WriteLine(typeof(A).IsAssignableFrom(typeof(B)));
Console.WriteLine(typeof(B).IsAssignableFrom(typeof(A)));"#
        ),
        &["True", "False"]
    );
}

#[test]
fn type_get_interfaces_includes_implemented_interfaces() {
    assert_eq!(
        run_csharp(
            r#"interface IFoo{}
class Foo:IFoo{}
bool has=System.Array.Exists(typeof(Foo).GetInterfaces(),t=>t==typeof(IFoo));
Console.WriteLine(has);"#
        ),
        &["True"]
    );
}

#[test]
fn type_base_type_reflects_inheritance_chain() {
    assert_eq!(
        run_csharp(
            r#"class A{} class B:A{} class C:B{}
Console.WriteLine(typeof(C).BaseType.Name);
Console.WriteLine(typeof(C).BaseType.BaseType.Name);"#
        ),
        &["B", "A"]
    );
}

#[test]
fn property_info_get_set_accessor_names() {
    assert_eq!(
        run_csharp(
            r#"class Model{public int Value{get;set;}}
var pi=typeof(Model).GetProperty("Value");
Console.WriteLine(pi.CanRead); Console.WriteLine(pi.CanWrite);"#
        ),
        &["True", "True"]
    );
}

#[test]
fn method_info_invoke_calls_instance_method() {
    assert_eq!(
        run_csharp(
            r#"class Adder{public int Add(int a,int b)=>a+b;}
var mi=typeof(Adder).GetMethod("Add");
var result=mi.Invoke(new Adder(),new object[]{3,4});
Console.WriteLine(result);"#
        ),
        &["7"]
    );
}

#[test]
fn field_info_get_and_set_value_on_instance() {
    assert_eq!(
        run_csharp(
            r#"class Box{public int V;}
var fi=typeof(Box).GetField("V");
var obj=new Box();
fi.SetValue(obj,55);
Console.WriteLine(fi.GetValue(obj));"#
        ),
        &["55"]
    );
}

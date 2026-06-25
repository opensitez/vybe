//! `record struct`: value semantics, `with`, equality, `ToString`.
use super::helpers::run_csharp;

#[test]
fn record_struct_with_expression_creates_new_copy() {
    assert_eq!(
        run_csharp(r#"record struct Point(int X,int Y);
var a=new Point(1,2);
var b=a with{X=99};
Console.WriteLine(a.X); Console.WriteLine(b.X);"#),
        &["1", "99"]
    );
}

#[test]
fn record_struct_equality_compares_by_value() {
    assert_eq!(
        run_csharp(r#"record struct Color(int R,int G,int B);
var c1=new Color(255,0,0); var c2=new Color(255,0,0);
Console.WriteLine(c1==c2);"#),
        &["True"]
    );
}

#[test]
fn record_struct_deconstruct_in_let_statement() {
    assert_eq!(
        run_csharp(r#"record struct Vec(int X,int Y);
var v=new Vec(3,4);
var(x,y)=v;
Console.WriteLine(x+y);"#),
        &["7"]
    );
}

#[test]
fn record_struct_to_string_includes_property_values() {
    assert_eq!(
        run_csharp(r#"record struct Tag(string Name);
Console.WriteLine(new Tag("admin").ToString().Contains("admin"));"#),
        &["True"]
    );
}

#[test]
fn record_struct_copy_is_independent_value_after_mutation() {
    assert_eq!(
        run_csharp(r#"record struct Count(int N);
var a=new Count(5);
var b=a;
b=b with{N=99};
Console.WriteLine(a.N);"#),
        &["5"]
    );
}

//! Method overloading resolution: parameter count, type specificity, and generics.
use super::helpers::run_csharp;

#[test]
fn overload_with_different_argument_count_dispatches_correctly() {
    assert_eq!(
        run_csharp(
            r#"string Desc(int a)=>"one";
string Desc(int a,int b)=>"two";
Console.WriteLine(Desc(1)); Console.WriteLine(Desc(1,2));"#
        ),
        &["one", "two"]
    );
}

#[test]
fn overload_on_type_picks_most_specific_match() {
    assert_eq!(
        run_csharp(
            r#"string Label(object o)=>"object";
string Label(string s)=>"string";
Console.WriteLine(Label("hi"));
Console.WriteLine(Label((object)"hi"));"#
        ),
        &["string", "object"]
    );
}

#[test]
fn overload_between_int_and_double_picks_exact_int_match() {
    assert_eq!(
        run_csharp(
            r#"string Kind(int n)=>"int";
string Kind(double d)=>"double";
Console.WriteLine(Kind(5));
Console.WriteLine(Kind(5.0));"#
        ),
        &["int", "double"]
    );
}

#[test]
fn generic_overload_less_specific_than_typed_overload() {
    assert_eq!(
        run_csharp(
            r#"string Foo<T>(T v)=>"generic";
string Foo(int v)=>"specific";
Console.WriteLine(Foo(1));
Console.WriteLine(Foo("x"));"#
        ),
        &["specific", "generic"]
    );
}

#[test]
fn overload_with_params_array_chosen_when_explicit_available() {
    assert_eq!(
        run_csharp(
            r#"string Sum(int a,int b)=>"two";
string Sum(params int[] ns)=>"params";
Console.WriteLine(Sum(1,2));
Console.WriteLine(Sum(1,2,3));"#
        ),
        &["two", "params"]
    );
}

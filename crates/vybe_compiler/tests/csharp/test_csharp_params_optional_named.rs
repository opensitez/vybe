//! `params`, optional, and named parameters.
use super::helpers::run_csharp;

#[test]
fn params_accepts_variable_number_of_arguments() {
    assert_eq!(
        run_csharp(r#"int Sum(params int[] ns){int s=0;foreach(var n in ns)s+=n;return s;}
Console.WriteLine(Sum(1,2,3,4,5));"#),
        &["15"]
    );
}

#[test]
fn params_can_be_called_with_zero_arguments() {
    assert_eq!(
        run_csharp(r#"int Sum(params int[] ns){int s=0;foreach(var n in ns)s+=n;return s;}
Console.WriteLine(Sum());"#),
        &["0"]
    );
}

#[test]
fn params_can_be_called_with_explicit_array() {
    assert_eq!(
        run_csharp(r#"int Sum(params int[] ns){int s=0;foreach(var n in ns)s+=n;return s;}
Console.WriteLine(Sum(new int[]{10,20}));"#),
        &["30"]
    );
}

#[test]
fn optional_parameter_uses_default_when_omitted() {
    assert_eq!(
        run_csharp(r#"string Greet(string name, string prefix="Hello") => prefix+" "+name;
Console.WriteLine(Greet("World"));"#),
        &["Hello World"]
    );
}

#[test]
fn optional_parameter_overridden_when_supplied() {
    assert_eq!(
        run_csharp(r#"string Greet(string name, string prefix="Hello") => prefix+" "+name;
Console.WriteLine(Greet("World","Hi"));"#),
        &["Hi World"]
    );
}

#[test]
fn named_argument_can_be_passed_out_of_order() {
    assert_eq!(
        run_csharp(r#"string Concat(string a, string b, string c) => a+b+c;
Console.WriteLine(Concat(c:"3",a:"1",b:"2"));"#),
        &["123"]
    );
}

#[test]
fn mix_of_positional_and_named_arguments() {
    assert_eq!(
        run_csharp(r#"int Sub(int x, int y) => x-y;
Console.WriteLine(Sub(10, y:3));"#),
        &["7"]
    );
}

#[test]
fn optional_with_null_default_allows_omission() {
    assert_eq!(
        run_csharp(r#"string Label(string text, string tag=null) => tag==null?text:$"[{tag}]{text}";
Console.WriteLine(Label("msg"));
Console.WriteLine(Label("msg","info"));"#),
        &["msg", "[info]msg"]
    );
}

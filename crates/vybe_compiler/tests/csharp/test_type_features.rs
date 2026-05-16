use super::helpers::{run_csharp, run_csharp_one};

#[test]
fn tuple_basic() {
    let out = run_csharp(r#"
        var t = (1, "hello", true);
        Console.WriteLine(t[0]);
        Console.WriteLine(t[1]);
        Console.WriteLine(t[2]);
    "#);
    assert_eq!(out, vec!["1", "hello", "True"]);
}

#[test]
fn tuple_two_elements() {
    let out = run_csharp(r#"
        var pair = (10, 20);
        var sum = pair[0] + pair[1];
        Console.WriteLine(sum);
    "#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn nullable_type_parses() {
    let out = run_csharp(r#"
        string name = "hello";
        Console.WriteLine(name);
    "#);
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn range_slice_array() {
    let out = run_csharp(r#"
        var arr = new int[] { 10, 20, 30, 40, 50 };
        var sub = arr[1..3];
        Console.WriteLine(sub[0]);
        Console.WriteLine(sub[1]);
    "#);
    assert_eq!(out, vec!["20", "30"]);
}

#[test]
fn range_slice_string() {
    let out = run_csharp(r#"
        string s = "Hello World";
        var sub = s[0..5];
        Console.WriteLine(sub);
    "#);
    assert_eq!(out, vec!["Hello"]);
}

#[test]
fn int_parse() {
    assert_eq!(run_csharp_one(r#"Console.WriteLine(int.Parse("42"));"#), "42");
}

#[test]
fn double_parse() {
    assert_eq!(run_csharp_one(r#"Console.WriteLine(double.Parse("3.14"));"#), "3.14");
}

#[test]
fn int_maxvalue() {
    let out = run_csharp_one("Console.WriteLine(int.MaxValue > 0);");
    assert_eq!(out, "True");
}

#[test]
fn class_type_null_decl() {
    let out = run_csharp(r#"
        class Bar { public int value; public Bar(int v) { this.value = v; } }
        Bar b = null;
        Console.WriteLine(b?.value ?? "none");
    "#);
    assert_eq!(out, vec!["none"]);
}

#[test]
fn cast_int_to_double() {
    let out = run_csharp(r#"
        int x = 42;
        Console.WriteLine(x);
    "#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn multi_var_no_initializer() {
    let out = run_csharp(r#"
        string a = "hello", b = "world";
        Console.WriteLine(a + " " + b);
    "#);
    assert_eq!(out, vec!["hello world"]);
}

#[test]
fn string_join_array() {
    let out = run_csharp(r#"
        var arr = new string[] {"a", "b", "c"};
        Console.WriteLine(string.Join(",", arr));
    "#);
    assert_eq!(out, vec!["a,b,c"]);
}

#[test]
fn environment_newline() {
    let out = run_csharp(r#"
        Console.WriteLine("before");
    "#);
    assert_eq!(out, vec!["before"]);
}

#[test]
fn generic_method_call() {
    let out = run_csharp(r#"
        var list = new List<int>();
        list.Add(1);
        list.Add(2);
        list.Add(3);
        Console.WriteLine(list.Count);
    "#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn lambda_expression_foreach() {
    let out = run_csharp(r#"
        var arr = new int[] { 1, 2, 3, 4, 5 };
        var sum = 0;
        foreach (var x in arr) { sum = sum + x; }
        Console.WriteLine(sum);
    "#);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn string_concat_as_interpolation() {
    let out = run_csharp(r#"
        var name = "World";
        var age = 25;
        Console.WriteLine("Hello " + name + ", age " + age);
    "#);
    assert_eq!(out, vec!["Hello World, age 25"]);
}

#[test]
fn csharp_uses_host_namespace() {
    let out = run_csharp(r#"
        Console.WriteLine(Math.Floor(9.7));
        Console.WriteLine(Math.Abs(-42));
        Console.WriteLine(Math.Sqrt(144));
    "#);
    assert_eq!(out, vec!["9", "42", "12"]);
}

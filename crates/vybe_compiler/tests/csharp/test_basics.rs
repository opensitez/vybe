use super::helpers::{run_csharp, run_csharp_one};

#[test]
fn hello_world() {
    let out = run_csharp(r#"Console.WriteLine("Hello, World!");"#);
    assert_eq!(out, vec!["Hello, World!"]);
}

#[test]
fn console_writeline_number() {
    assert_eq!(run_csharp_one("Console.WriteLine(42);"), "42");
}

#[test]
fn console_writeline_bool() {
    assert_eq!(run_csharp_one("Console.WriteLine(true);"), "True");
}

#[test]
fn var_declaration() {
    let out = run_csharp(
        r#"
        var x = 10;
        var y = 20;
        Console.WriteLine(x + y);
    "#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn typed_declaration() {
    let out = run_csharp(
        r#"
        int x = 5;
        double y = 3.14;
        Console.WriteLine(x);
        Console.WriteLine(y);
    "#,
    );
    assert_eq!(out, vec!["5", "3.14"]);
}

#[test]
fn string_concat() {
    assert_eq!(
        run_csharp_one(r#"Console.WriteLine("hello" + " " + "world");"#),
        "hello world"
    );
}

#[test]
fn arithmetic() {
    assert_eq!(run_csharp_one("Console.WriteLine(2 + 3 * 4);"), "14");
}

#[test]
fn comparison() {
    assert_eq!(run_csharp_one("Console.WriteLine(5 > 3);"), "True");
    assert_eq!(run_csharp_one("Console.WriteLine(5 < 3);"), "False");
}

#[test]
fn compound_assignment() {
    let out = run_csharp(
        r#"
        var x = 10;
        x += 5;
        x -= 3;
        x *= 2;
        Console.WriteLine(x);
    "#,
    );
    assert_eq!(out, vec!["24"]);
}

#[test]
fn boolean_and_or() {
    let out = run_csharp(
        r#"
        Console.WriteLine(true && false);
        Console.WriteLine(true || false);
        Console.WriteLine(!true);
    "#,
    );
    assert_eq!(out, vec!["False", "True", "False"]);
}

#[test]
fn postfix_increment() {
    let out = run_csharp(
        r#"
        var x = 5;
        x++;
        x++;
        Console.WriteLine(x);
    "#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn prefix_decrement() {
    let out = run_csharp(
        r#"
        var x = 10;
        --x;
        --x;
        Console.WriteLine(x);
    "#,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn multi_var_declaration() {
    let out = run_csharp(
        r#"
        int a = 1, b = 2;
        Console.WriteLine(a + b);
    "#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn multi_var_three_variables() {
    let out = run_csharp(
        r#"
        int x = 10, y = 20, z = 30;
        Console.WriteLine(x + y + z);
    "#,
    );
    assert_eq!(out, vec!["60"]);
}

#[test]
fn ternary_expression() {
    assert_eq!(
        run_csharp_one("Console.WriteLine(5 > 3 ? \"yes\" : \"no\");"),
        "yes"
    );
}

#[test]
fn null_coalescing() {
    assert_eq!(
        run_csharp_one(r#"string s = null; Console.WriteLine(s ?? "default");"#),
        "default"
    );
}

#[test]
fn typeof_expression() {
    // .NET: `Console.WriteLine(typeof(int))` calls Type.ToString() →
    // returns FullName ("System.Int32"), not the C# alias.
    assert_eq!(
        run_csharp_one(r#"Console.WriteLine(typeof(int));"#),
        "System.Int32"
    );
}

#[test]
fn nameof_expression() {
    let out = run_csharp(
        r#"
        var myVar = 42;
        Console.WriteLine(nameof(myVar));
    "#,
    );
    assert_eq!(out, vec!["myVar"]);
}

#[test]
fn default_int() {
    assert_eq!(run_csharp_one("Console.WriteLine(default(int));"), "0");
}

#[test]
fn default_bool() {
    assert_eq!(run_csharp_one("Console.WriteLine(default(bool));"), "False");
}

#[test]
fn string_plus_number() {
    assert_eq!(
        run_csharp_one(r#"Console.WriteLine("Value: " + 42);"#),
        "Value: 42"
    );
}

#[test]
fn array_copy_runtime() {
    // C# `Array.Copy(src, dst, count)` — copies first 3 elems from src into dst
    let out = run_csharp_one(
        r#"
int[] src = new int[] { 10, 20, 30, 40 };
int[] dst = new int[] { 0, 0, 0, 0 };
Array.Copy(src, dst, 3);
Console.WriteLine(dst[0] + dst[1] + dst[2] + dst[3]);
"#,
    );
    assert_eq!(out, "60");
}

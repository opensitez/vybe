use super::helpers::run_csharp;

// ═══════════════════════════════════════════════════════════
// C#: Lambdas, delegates, closures, functional patterns
// ═══════════════════════════════════════════════════════════

#[test]
fn lambda_expression() {
    let out = run_csharp(
        r#"
var double_it = (int x) => x * 2;
Console.WriteLine(double_it(5));
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn lambda_block_body() {
    let out = run_csharp(
        r#"
var add = (int a, int b) => {
    return a + b;
};
Console.WriteLine(add(3, 4));
"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn lambda_closure() {
    let out = run_csharp(
        r#"
int counter = 0;
var inc = () => { counter++; };
inc();
inc();
inc();
Console.WriteLine(counter);
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn lambda_as_callback() {
    let out = run_csharp(
        r#"
int Apply(Func<int, int> fn, int x) {
    return fn(x);
}
Console.WriteLine(Apply(x => x * x, 5));
"#,
    );
    assert_eq!(out, vec!["25"]);
}

#[test]
fn lambda_with_list_foreach() {
    let out = run_csharp(
        r#"
using System.Collections.Generic;
var items = new List<int>();
items.Add(1);
items.Add(2);
items.Add(3);
items.ForEach(x => Console.WriteLine(x));
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn function_returning_function() {
    let out = run_csharp(
        r#"
Func<int, int> Multiplier(int factor) {
    return x => x * factor;
}
var triple = Multiplier(3);
Console.WriteLine(triple(7));
"#,
    );
    assert_eq!(out, vec!["21"]);
}

#[test]
fn higher_order_function() {
    let out = run_csharp(
        r#"
Func<int, int> square = x => x * x;
Func<int, int> negate = x => -x;
Console.WriteLine(square(5));
Console.WriteLine(negate(5));
"#,
    );
    assert_eq!(out, vec!["25", "-5"]);
}

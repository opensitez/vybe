//! `with` expressions on records produce new instances with specified fields changed.
use super::helpers::run_csharp;

#[test]
fn with_expression_creates_new_record_preserving_unchanged_fields() {
    assert_eq!(
        run_csharp(
            r#"
record Point(int X, int Y);
var origin = new Point(1, 2);
var moved = origin with { X = 10 };
Console.WriteLine(moved.X);
Console.WriteLine(moved.Y);
"#
        ),
        &["10", "2"]
    );
}

#[test]
fn with_expression_original_record_is_not_mutated() {
    assert_eq!(
        run_csharp(
            r#"
record Point(int X, int Y);
var origin = new Point(1, 2);
var moved = origin with { X = 10 };
Console.WriteLine(origin.X);
"#
        ),
        &["1"]
    );
}

#[test]
fn with_expression_changing_two_properties_at_once() {
    assert_eq!(
        run_csharp(
            r#"
record Person(string Name, int Age);
var p = new Person("Ada", 30);
var updated = p with { Name = "Grace", Age = 31 };
Console.WriteLine(updated.Name);
Console.WriteLine(updated.Age);
"#
        ),
        &["Grace", "31"]
    );
}

#[test]
fn with_expression_chained_produces_independent_copies() {
    assert_eq!(
        run_csharp(
            r#"
record Box(int Width, int Height);
var a = new Box(1, 1);
var b = a with { Width = 2 };
var c = b with { Height = 3 };
Console.WriteLine(a.Width);
Console.WriteLine(c.Width);
Console.WriteLine(c.Height);
"#
        ),
        &["1", "2", "3"]
    );
}

#[test]
fn with_expression_on_record_with_init_property() {
    assert_eq!(
        run_csharp(
            r#"
record Config { public string Host { get; init; } public int Port { get; init; } }
var base_ = new Config { Host = "localhost", Port = 80 };
var prod = base_ with { Port = 443 };
Console.WriteLine(prod.Host);
Console.WriteLine(prod.Port);
"#
        ),
        &["localhost", "443"]
    );
}

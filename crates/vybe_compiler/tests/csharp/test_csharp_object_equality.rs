//! `Equals`, `GetHashCode`, `==`, and `ReferenceEquals` contracts.
use super::helpers::run_csharp;

#[test]
fn reference_equals_returns_false_for_two_distinct_object_instances() {
    assert_eq!(
        run_csharp(
            r#"
var a = new object();
var b = new object();
Console.WriteLine(object.ReferenceEquals(a, b));
"#
        ),
        &["False"]
    );
}

#[test]
fn reference_equals_returns_true_for_same_reference() {
    assert_eq!(
        run_csharp(
            r#"
var a = new object();
var b = a;
Console.WriteLine(object.ReferenceEquals(a, b));
"#
        ),
        &["True"]
    );
}

#[test]
fn string_equals_compares_content_not_reference() {
    assert_eq!(
        run_csharp(
            r#"
string a = new string(new char[] { 'h', 'i' });
string b = new string(new char[] { 'h', 'i' });
Console.WriteLine(a.Equals(b));
"#
        ),
        &["True"]
    );
}

#[test]
fn overridden_equals_on_class_reflects_value_semantics() {
    assert_eq!(
        run_csharp(
            r#"
class Money {
    public int Amount;
    public override bool Equals(object obj) =>
        obj is Money m && m.Amount == Amount;
    public override int GetHashCode() => Amount;
}
var x = new Money { Amount = 5 };
var y = new Money { Amount = 5 };
Console.WriteLine(x.Equals(y));
Console.WriteLine(object.ReferenceEquals(x, y));
"#
        ),
        &["True", "False"]
    );
}

#[test]
fn equal_hash_codes_required_for_equal_objects() {
    assert_eq!(
        run_csharp(
            r#"
class Key {
    public int Id;
    public override bool Equals(object obj) => obj is Key k && k.Id == Id;
    public override int GetHashCode() => Id;
}
var x = new Key { Id = 7 };
var y = new Key { Id = 7 };
Console.WriteLine(x.GetHashCode() == y.GetHashCode());
"#
        ),
        &["True"]
    );
}

#[test]
fn null_equals_null_returns_true_via_static_method() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(object.Equals(null, null));"#),
        &["True"]
    );
}

#[test]
fn record_equality_compares_all_positional_properties() {
    assert_eq!(
        run_csharp(
            r#"
record Point(int X, int Y);
var a = new Point(1, 2);
var b = new Point(1, 2);
var c = new Point(1, 3);
Console.WriteLine(a.Equals(b));
Console.WriteLine(a.Equals(c));
"#
        ),
        &["True", "False"]
    );
}

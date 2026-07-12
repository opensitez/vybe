//! Deconstruction patterns: positional, property, list, and nested patterns.
use super::helpers::run_csharp;

#[test]
fn positional_pattern_matches_deconstructed_record_fields() {
    assert_eq!(
        run_csharp(
            r#"
record Point(int X, int Y);
object obj = new Point(0, 5);
if (obj is Point(0, var y)) Console.WriteLine(y);
else Console.WriteLine(-1);
"#
        ),
        &["5"]
    );
}

#[test]
fn property_pattern_extracts_nested_property_value() {
    assert_eq!(
        run_csharp(
            r#"
class Order { public int Amount; public bool IsPaid; }
object o = new Order { Amount = 100, IsPaid = true };
var label = o switch {
    Order { IsPaid: true, Amount: > 50 } => "big paid",
    Order { IsPaid: true }               => "small paid",
    _                                    => "unpaid"
};
Console.WriteLine(label);
"#
        ),
        &["big paid"]
    );
}

#[test]
fn list_pattern_matches_exact_element_sequence() {
    assert_eq!(
        run_csharp(
            r#"
int[] data = { 1, 2, 3 };
if (data is [1, 2, 3]) Console.WriteLine("exact");
else Console.WriteLine("no");
"#
        ),
        &["exact"]
    );
}

#[test]
fn list_pattern_with_slice_matches_prefix_and_suffix() {
    assert_eq!(
        run_csharp(
            r#"
int[] data = { 1, 2, 3, 4, 5 };
if (data is [1, .., 5]) Console.WriteLine("bookended");
else Console.WriteLine("no");
"#
        ),
        &["bookended"]
    );
}

#[test]
fn nested_property_pattern_inspects_inner_object_fields() {
    assert_eq!(
        run_csharp(
            r#"
class Address { public string City; }
class Person { public Address Home; }
object p = new Person { Home = new Address { City = "Paris" } };
if (p is Person { Home: { City: "Paris" } }) Console.WriteLine("Paris");
else Console.WriteLine("elsewhere");
"#
        ),
        &["Paris"]
    );
}

#[test]
fn var_pattern_binds_matched_value_regardless_of_type() {
    assert_eq!(
        run_csharp(
            r#"
object value = 42;
if (value is var captured) Console.WriteLine(captured);
"#
        ),
        &["42"]
    );
}

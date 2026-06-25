//! Equality contracts: reference identity vs value equality for strings,
//! structs, and boxed values.
use super::helpers::run_csharp;

#[test]
fn string_equality_compares_character_sequence_not_reference_identity() {
    assert_eq!(
        run_csharp(
            r#"
string a = new string(new char[] { 'h', 'i' });
string b = new string(new char[] { 'h', 'i' });
Console.WriteLine(a == b);
Console.WriteLine(object.ReferenceEquals(a, b));
"#
        ),
        &["True", "False"]
    );
}

#[test]
fn struct_equality_compares_field_values_when_equals_overridden() {
    assert_eq!(
        run_csharp(
            r#"
struct Point {
    public int X;
    public int Y;
    public bool Equals(Point other) { return X == other.X && Y == other.Y; }
}
var left = new Point { X = 2, Y = 3 };
var right = new Point { X = 2, Y = 3 };
Console.WriteLine(left.Equals(right));
"#
        ),
        &["True"]
    );
}

#[test]
fn boxed_value_types_with_same_numeric_value_compare_equal_with_equals() {
    assert_eq!(
        run_csharp(
            r#"
object left = 42;
object right = 42;
Console.WriteLine(left.Equals(right));
"#
        ),
        &["True"]
    );
}

#[test]
fn dictionary_with_custom_ignore_case_comparer_treats_keys_as_equivalent() {
    assert_eq!(
        run_csharp(
            r#"
using System.Collections.Generic;
var map = new Dictionary<string, int>(System.StringComparer.OrdinalIgnoreCase);
map["User"] = 1;
Console.WriteLine(map.ContainsKey("user"));
Console.WriteLine(map["USER"]);
"#
        ),
        &["True", "1"]
    );
}

#[test]
fn list_reference_equality_is_false_for_distinct_instances_with_same_contents() {
    assert_eq!(
        run_csharp(
            r#"
using System.Collections.Generic;
using System.Linq;
var left = new List<int> { 1, 2 };
var right = new List<int> { 1, 2 };
Console.WriteLine(left == right);
Console.WriteLine(left.SequenceEqual(right));
"#
        ),
        &["False", "True"]
    );
}

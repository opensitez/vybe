//! Struct assignment copies value; fields on one copy do not mutate the other.
use super::helpers::run_csharp;

#[test]
fn struct_field_mutation_on_copy_leaves_original_unchanged() {
    assert_eq!(
        run_csharp(
            r#"
struct Point { public int X; }
var left = new Point { X = 1 };
var right = left;
right.X = 9;
Console.WriteLine(left.X);
Console.WriteLine(right.X);
"#
        ),
        &["1", "9"]
    );
}

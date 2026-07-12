//! Target-typed `new` infers type from the variable being assigned.
use super::helpers::run_csharp;

#[test]
fn target_typed_new_creates_list_without_repeating_type_arguments() {
    assert_eq!(
        run_csharp(
            r#"
System.Collections.Generic.List<int> values = new();
values.Add(7);
Console.WriteLine(values[0]);
"#
        ),
        &["7"]
    );
}
